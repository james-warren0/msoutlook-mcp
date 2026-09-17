/**
 * Mail API — email CRUD, search, and folder operations.
 */

import { readFile } from 'node:fs/promises';
import { basename, extname } from 'node:path';
import { owaGet, owaPost, owaPatch, owaDelete } from './client.js';

function getFromMailbox<T>(path: string, params?: Record<string, string>, mailbox?: string): Promise<T> {
  if (mailbox) return owaGet<T>(path, params, mailbox);
  return owaGet<T>(path, params);
}

function postToMailbox<T>(path: string, body: unknown, mailbox?: string): Promise<T> {
  return mailbox ? owaPost<T>(path, body, mailbox) : owaPost<T>(path, body);
}

function patchInMailbox<T>(path: string, body: unknown, mailbox?: string): Promise<T> {
  return mailbox ? owaPatch<T>(path, body, mailbox) : owaPatch<T>(path, body);
}

function deleteFromMailbox(path: string, mailbox?: string): Promise<void> {
  return mailbox ? owaDelete(path, mailbox) : owaDelete(path);
}

// ─────────────────────────────────────────────────────────────────────────────
// Body formatting
// ─────────────────────────────────────────────────────────────────────────────

const HTML_TAG_RE = /<\/?[a-z][^>]*>/i;

/** True when the string already carries HTML markup we should leave untouched. */
export function looksLikeHtml(s: string): boolean {
  return HTML_TAG_RE.test(s);
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Normalise an outgoing body to HTML so line breaks survive.
 *
 * Outlook renders reply/forward comments and HTML bodies as HTML, where a raw
 * "\n" collapses and the whole message arrives as one block. To stop that:
 * - if the body already contains HTML markup, pass it through untouched, but
 * - if it is plain text, escape it and turn newlines into <br> (a blank line,
 *   i.e. a double newline, becomes <br><br> so paragraphs keep their gap).
 */
export function toHtmlBody(body: string): string {
  if (looksLikeHtml(body)) return body;
  return escapeHtml(body).replace(/\r\n?/g, '\n').replace(/\n/g, '<br>');
}

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface EmailAddress {
  Name: string;
  Address: string;
}

export interface Recipient {
  EmailAddress: EmailAddress;
}

export interface Message {
  Id: string;
  Subject: string;
  BodyPreview?: string;
  Body?: { ContentType: string; Content: string };
  From?: Recipient;
  Sender?: Recipient;
  ToRecipients?: Recipient[];
  CcRecipients?: Recipient[];
  BccRecipients?: Recipient[];
  ReceivedDateTime?: string;
  SentDateTime?: string;
  IsRead?: boolean;
  HasAttachments?: boolean;
  Importance?: string;
  Flag?: { FlagStatus: string };
  Categories?: string[];
  IsDraft?: boolean;
  ConversationId?: string;
  ParentFolderId?: string;
  WebLink?: string;
  Attachments?: Attachment[];
  InternetMessageId?: string;
  InternetMessageHeaders?: InternetMessageHeader[];
}

export interface InternetMessageHeader {
  Name: string;
  Value: string;
}

export interface Attachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  IsInline: boolean;
}

export interface MailFolder {
  Id: string;
  DisplayName: string;
  UnreadItemCount: number;
  TotalItemCount: number;
  ChildFolderCount: number;
}

export interface ODataResponse<T> {
  value: T[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// List messages
// ─────────────────────────────────────────────────────────────────────────────

export interface ListMessagesOptions {
  /** Shared/delegated mailbox UPN. Defaults to the signed-in user's mailbox. */
  mailbox?: string;
  folder?: string;
  top?: number;
  skip?: number;
  filter?: string;
  select?: string[];
  orderBy?: string;
  search?: string;
}

export async function listMessages(opts: ListMessagesOptions = {}): Promise<Message[]> {
  const folder = opts.folder ?? 'Inbox';
  const params: Record<string, string> = {
    '$top': String(opts.top ?? 20),
    '$select': (opts.select ?? [
      'Id', 'Subject', 'BodyPreview', 'From', 'ToRecipients', 'ReceivedDateTime',
      'IsRead', 'HasAttachments', 'Importance', 'Flag', 'ConversationId', 'WebLink',
    ]).join(','),
  };

  if (opts.skip) params['$skip'] = String(opts.skip);

  // $search cannot be combined with $orderby or $filter — the API returns
  // 400 SearchWithOrderBy. Search results are relevance-ranked, so we drop
  // ordering when searching.
  if (opts.search) {
    params['$search'] = `"${opts.search}"`;
  } else {
    params['$orderby'] = opts.orderBy ?? 'ReceivedDateTime desc';
    if (opts.filter) params['$filter'] = opts.filter;
  }

  const res = await getFromMailbox<ODataResponse<Message>>(`/MailFolders/${folder}/messages`, params, opts.mailbox);
  return res.value;
}

// ─────────────────────────────────────────────────────────────────────────────
// Get message
// ─────────────────────────────────────────────────────────────────────────────

export async function getMessage(id: string, includeAttachments = false, mailbox?: string): Promise<Message> {
  const params = includeAttachments ? { '$expand': 'Attachments' } : undefined;
  return getFromMailbox<Message>(`/messages/${id}`, params, mailbox);
}

/** Read the RFC-style transport headers Outlook exposes for a message. */
export async function getMessageHeaders(id: string, mailbox?: string): Promise<Message> {
  return getFromMailbox<Message>(`/messages/${id}`, {
    '$select': 'Id,Subject,InternetMessageId,InternetMessageHeaders',
  }, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Send email
// ─────────────────────────────────────────────────────────────────────────────

/** A file attachment ready to send (content already base64 encoded). */
export interface OutgoingAttachment {
  name: string;
  contentType: string;
  contentBytes: string;
}

function toOwaAttachment(a: OutgoingAttachment): Record<string, unknown> {
  return {
    '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
    Name: a.name,
    ContentType: a.contentType,
    ContentBytes: a.contentBytes,
  };
}

const CONTENT_TYPES: Record<string, string> = {
  '.pdf': 'application/pdf', '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
  '.gif': 'image/gif', '.txt': 'text/plain', '.csv': 'text/csv', '.html': 'text/html',
  '.doc': 'application/msword', '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.zip': 'application/zip', '.json': 'application/json', '.ics': 'text/calendar',
};

/** Read a file from disk into an outgoing attachment (base64 encoded). */
export async function fileToAttachment(path: string): Promise<OutgoingAttachment> {
  const data = await readFile(path);
  return {
    name: basename(path),
    contentType: CONTENT_TYPES[extname(path).toLowerCase()] ?? 'application/octet-stream',
    contentBytes: data.toString('base64'),
  };
}

export interface SendEmailOptions {
  /** Shared/delegated mailbox UPN. Defaults to the signed-in user's mailbox. */
  mailbox?: string;
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  bodyType?: 'Text' | 'HTML';
  importance?: 'Low' | 'Normal' | 'High';
  attachments?: OutgoingAttachment[];
  saveToSentItems?: boolean;
}

export async function sendEmail(opts: SendEmailOptions): Promise<void> {
  const toRecipients = opts.to.map(addr => ({
    EmailAddress: { Address: addr },
  }));
  const ccRecipients = opts.cc?.map(addr => ({ EmailAddress: { Address: addr } }));
  const bccRecipients = opts.bcc?.map(addr => ({ EmailAddress: { Address: addr } }));

  const bodyType = opts.bodyType ?? 'HTML';
  const content = bodyType === 'HTML' ? toHtmlBody(opts.body) : opts.body;

  await postToMailbox('/sendmail', {
    Message: {
      Subject: opts.subject,
      Body: { ContentType: bodyType, Content: content },
      ToRecipients: toRecipients,
      ...(ccRecipients?.length ? { CcRecipients: ccRecipients } : {}),
      ...(bccRecipients?.length ? { BccRecipients: bccRecipients } : {}),
      Importance: opts.importance ?? 'Normal',
      ...(opts.attachments?.length ? { Attachments: opts.attachments.map(toOwaAttachment) } : {}),
    },
    SaveToSentItems: opts.saveToSentItems !== false,
  }, opts.mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Create draft
// ─────────────────────────────────────────────────────────────────────────────

export async function createDraft(opts: Omit<SendEmailOptions, 'saveToSentItems'>): Promise<Message> {
  const toRecipients = opts.to.map(addr => ({ EmailAddress: { Address: addr } }));
  const ccRecipients = opts.cc?.map(addr => ({ EmailAddress: { Address: addr } }));
  const bccRecipients = opts.bcc?.map(addr => ({ EmailAddress: { Address: addr } }));

  const bodyType = opts.bodyType ?? 'HTML';
  const content = bodyType === 'HTML' ? toHtmlBody(opts.body) : opts.body;

  return postToMailbox<Message>('/messages', {
    Subject: opts.subject,
    Body: { ContentType: bodyType, Content: content },
    ToRecipients: toRecipients,
    ...(ccRecipients?.length ? { CcRecipients: ccRecipients } : {}),
    ...(bccRecipients?.length ? { BccRecipients: bccRecipients } : {}),
    Importance: opts.importance ?? 'Normal',
    ...(opts.attachments?.length ? { Attachments: opts.attachments.map(toOwaAttachment) } : {}),
  }, opts.mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Reply
// ─────────────────────────────────────────────────────────────────────────────

export async function replyToMessage(id: string, body: string, replyAll = false, mailbox?: string): Promise<void> {
  const action = replyAll ? 'replyall' : 'reply';
  await postToMailbox(`/messages/${id}/${action}`, {
    Comment: toHtmlBody(body),
  }, mailbox);
}

/**
 * Create a reply (or reply-all) as a draft instead of sending it.
 *
 * Returns the new draft message, which already has the recipients and the
 * quoted original prefilled, with the supplied body inserted above the quote.
 * The caller can then review or edit it in Outlook and send it later with
 * sendDraft (outlook_send_draft).
 */
export async function createReplyDraft(id: string, body: string, replyAll = false, mailbox?: string): Promise<Message> {
  const action = replyAll ? 'createreplyall' : 'createreply';
  return postToMailbox<Message>(`/messages/${id}/${action}`, {
    Comment: toHtmlBody(body),
  }, mailbox);
}

/**
 * Create a forward as a draft instead of sending it. Returns the draft message
 * with the quoted original prefilled; recipients and any final edits can be set
 * in Outlook (or via a later update) before sending with sendDraft.
 */
export async function createForwardDraft(id: string, body = '', to?: string[], mailbox?: string): Promise<Message> {
  const payload: Record<string, unknown> = { Comment: body ? toHtmlBody(body) : '' };
  if (to?.length) {
    payload.ToRecipients = to.map(addr => ({ EmailAddress: { Address: addr } }));
  }
  return postToMailbox<Message>(`/messages/${id}/createforward`, payload, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Forward
// ─────────────────────────────────────────────────────────────────────────────

export async function forwardMessage(id: string, to: string[], comment?: string, mailbox?: string): Promise<void> {
  await postToMailbox(`/messages/${id}/forward`, {
    ToRecipients: to.map(addr => ({ EmailAddress: { Address: addr } })),
    Comment: comment ? toHtmlBody(comment) : '',
  }, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Mark read/unread
// ─────────────────────────────────────────────────────────────────────────────

export async function markMessageRead(id: string, isRead = true, mailbox?: string): Promise<void> {
  await patchInMailbox(`/messages/${id}`, { IsRead: isRead }, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Flag
// ─────────────────────────────────────────────────────────────────────────────

export async function flagMessage(id: string, status: 'Flagged' | 'Complete' | 'NotFlagged' = 'Flagged', mailbox?: string): Promise<void> {
  await patchInMailbox(`/messages/${id}`, { Flag: { FlagStatus: status } }, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Move
// ─────────────────────────────────────────────────────────────────────────────

export async function moveMessage(id: string, destinationFolderId: string, mailbox?: string): Promise<Message> {
  return postToMailbox<Message>(`/messages/${id}/move`, {
    DestinationId: destinationFolderId,
  }, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Delete
// ─────────────────────────────────────────────────────────────────────────────

export async function deleteMessage(id: string, mailbox?: string): Promise<void> {
  await deleteFromMailbox(`/messages/${id}`, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Search
// ─────────────────────────────────────────────────────────────────────────────

export interface SearchMessagesOptions {
  /** Shared/delegated mailbox UPN. Defaults to the signed-in user's mailbox. */
  mailbox?: string;
  /** Keyword query. Optional: omit it to list everything in a date range. */
  query?: string;
  top?: number;
  folder?: string;
  /** Inclusive lower bound on received date (ISO date or datetime). */
  startDate?: string;
  /** Inclusive upper bound on received date (ISO date or datetime). */
  endDate?: string;
  /** Opaque page token from a previous result's nextSkipToken. */
  skipToken?: string;
}

export interface SearchMessagesResult {
  messages: Message[];
  /** Pass this back as skipToken to fetch the next page; absent when there are no more results. */
  nextSkipToken?: string;
}

/** Reduce an ISO date or datetime to the YYYY-MM-DD form KQL expects. */
function toKqlDate(value: string): string {
  return value.slice(0, 10);
}

/**
 * Search messages with optional date range and pagination.
 *
 * The OWA REST API rejects $skip, $filter and $orderby alongside $search, so:
 * - date range is expressed as KQL constraints inside the search string
 *   (received>=START AND received<=END), and
 * - pagination uses the $skiptoken returned in @odata.nextLink.
 */
export async function searchMessages(opts: SearchMessagesOptions): Promise<SearchMessagesResult> {
  const folder = opts.folder ?? 'Inbox';

  const parts: string[] = [];
  const keyword = opts.query?.trim();
  if (keyword) parts.push(keyword);
  if (opts.startDate) parts.push(`received>=${toKqlDate(opts.startDate)}`);
  if (opts.endDate) parts.push(`received<=${toKqlDate(opts.endDate)}`);
  if (parts.length === 0) {
    throw new Error('searchMessages needs a query and/or a date range (startDate or endDate).');
  }
  const kql = parts.join(' AND ');

  const params: Record<string, string> = {
    '$search': `"${kql}"`,
    '$top': String(opts.top ?? 20),
    '$select': 'Id,Subject,BodyPreview,From,ReceivedDateTime,IsRead,WebLink',
  };
  if (opts.skipToken) params['$skiptoken'] = opts.skipToken;

  const res = await getFromMailbox<ODataResponse<Message>>(`/MailFolders/${folder}/messages`, params, opts.mailbox);

  let nextSkipToken: string | undefined;
  const nextLink = res['@odata.nextLink'];
  if (nextLink) {
    try {
      nextSkipToken = new URL(nextLink).searchParams.get('$skiptoken') ?? undefined;
    } catch {
      nextSkipToken = undefined;
    }
  }

  return { messages: res.value, nextSkipToken };
}

// ─────────────────────────────────────────────────────────────────────────────
// Folders
// ─────────────────────────────────────────────────────────────────────────────

export async function listFolders(mailbox?: string): Promise<MailFolder[]> {
  const res = await getFromMailbox<ODataResponse<MailFolder>>('/mailfolders', {
    '$top': '50',
    '$select': 'Id,DisplayName,UnreadItemCount,TotalItemCount,ChildFolderCount',
  }, mailbox);
  return res.value;
}

export async function getFolder(idOrName: string, mailbox?: string): Promise<MailFolder> {
  return mailbox
    ? owaGet<MailFolder>(`/mailfolders/${idOrName}`, undefined, mailbox)
    : owaGet<MailFolder>(`/mailfolders/${idOrName}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Unread count
// ─────────────────────────────────────────────────────────────────────────────

export async function getUnreadMessages(top = 10, mailbox?: string): Promise<Message[]> {
  return listMessages({
    mailbox,
    filter: 'IsRead eq false',
    top,
    orderBy: 'ReceivedDateTime desc',
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Send draft
// ─────────────────────────────────────────────────────────────────────────────

export async function sendDraft(draftId: string, mailbox?: string): Promise<void> {
  await postToMailbox(`/messages/${draftId}/send`, {}, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Update a draft (recipients, subject, body)
// ─────────────────────────────────────────────────────────────────────────────

export interface UpdateDraftOptions {
  /** Shared/delegated mailbox UPN. Defaults to the signed-in user's mailbox. */
  mailbox?: string;
  subject?: string;
  body?: string;
  bodyType?: 'Text' | 'HTML';
  to?: string[];
  cc?: string[];
  bcc?: string[];
  importance?: 'Low' | 'Normal' | 'High';
}

export async function updateDraft(id: string, opts: UpdateDraftOptions): Promise<Message> {
  const patch: Record<string, unknown> = {};
  if (opts.subject !== undefined) patch.Subject = opts.subject;
  if (opts.body !== undefined) {
    const bodyType = opts.bodyType ?? 'HTML';
    patch.Body = { ContentType: bodyType, Content: bodyType === 'HTML' ? toHtmlBody(opts.body) : opts.body };
  }
  if (opts.to) patch.ToRecipients = opts.to.map(addr => ({ EmailAddress: { Address: addr } }));
  if (opts.cc) patch.CcRecipients = opts.cc.map(addr => ({ EmailAddress: { Address: addr } }));
  if (opts.bcc) patch.BccRecipients = opts.bcc.map(addr => ({ EmailAddress: { Address: addr } }));
  if (opts.importance) patch.Importance = opts.importance;
  return patchInMailbox<Message>(`/messages/${id}`, patch, opts.mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Attachments
// ─────────────────────────────────────────────────────────────────────────────

export interface FileAttachmentContent {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  ContentBytes?: string;
}

export async function listAttachments(messageId: string, mailbox?: string): Promise<Attachment[]> {
  const res = await getFromMailbox<ODataResponse<Attachment>>(`/messages/${messageId}/attachments`, {
    '$select': 'Id,Name,ContentType,Size,IsInline',
  }, mailbox);
  return res.value;
}

/** Fetch one attachment including its base64 content bytes. */
export async function getAttachmentContent(messageId: string, attachmentId: string, mailbox?: string): Promise<FileAttachmentContent> {
  return mailbox
    ? owaGet<FileAttachmentContent>(`/messages/${messageId}/attachments/${attachmentId}`, undefined, mailbox)
    : owaGet<FileAttachmentContent>(`/messages/${messageId}/attachments/${attachmentId}`);
}

/** Add a file attachment to an existing message/draft. */
export async function addAttachment(messageId: string, attachment: OutgoingAttachment, mailbox?: string): Promise<void> {
  await postToMailbox(`/messages/${messageId}/attachments`, toOwaAttachment(attachment), mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Conversation (thread)
// ─────────────────────────────────────────────────────────────────────────────

export async function getConversation(conversationId: string, top = 50, mailbox?: string): Promise<Message[]> {
  // Filtering by ConversationId cannot be combined with $orderby (the API
  // rejects it as an inefficient filter), so we sort by received date here.
  const res = await getFromMailbox<ODataResponse<Message>>('/messages', {
    '$filter': `ConversationId eq '${conversationId.replace(/'/g, "''")}'`,
    '$top': String(top),
    '$select': 'Id,Subject,BodyPreview,From,ToRecipients,ReceivedDateTime,IsRead,HasAttachments,ConversationId,WebLink',
  }, mailbox);
  return res.value.sort((a, b) => (a.ReceivedDateTime ?? '').localeCompare(b.ReceivedDateTime ?? ''));
}

// ─────────────────────────────────────────────────────────────────────────────
// Categories
// ─────────────────────────────────────────────────────────────────────────────

export async function setCategories(messageId: string, categories: string[], mailbox?: string): Promise<Message> {
  return patchInMailbox<Message>(`/messages/${messageId}`, { Categories: categories }, mailbox);
}

// ─────────────────────────────────────────────────────────────────────────────
// Folder management
// ─────────────────────────────────────────────────────────────────────────────

export async function createFolder(displayName: string, parentFolderId?: string, mailbox?: string): Promise<MailFolder> {
  const path = parentFolderId ? `/mailfolders/${parentFolderId}/childfolders` : '/mailfolders';
  return postToMailbox<MailFolder>(path, { DisplayName: displayName }, mailbox);
}

export async function renameFolder(id: string, displayName: string, mailbox?: string): Promise<MailFolder> {
  return patchInMailbox<MailFolder>(`/mailfolders/${id}`, { DisplayName: displayName }, mailbox);
}

export async function deleteFolder(id: string, mailbox?: string): Promise<void> {
  await deleteFromMailbox(`/mailfolders/${id}`, mailbox);
}
