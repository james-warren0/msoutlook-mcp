/**
 * Base API client for Outlook Web API calls.
 *
 * Uses the OWA REST API v2, targeting either the signed-in user's `/me`
 * resource or a delegated mailbox's `/users/{upn}` resource. This is the same
 * API the Outlook mobile and web clients use internally. The token from the
 * OWA session grants access to the scopes already available to that account.
 */

import { getOwaToken, getGraphToken } from '../auth/index.js';
import { OWA_REST_V2, GRAPH_BASE, OWA_BASE } from '../constants.js';
import { getBearerHeaders, parseResponse, fetchWithRetry } from '../utils/http.js';

// ─────────────────────────────────────────────────────────────────────────────
// OWA REST API client
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Build the mailbox resource prefix used by Outlook REST.
 *
 * Calls target the signed-in user's mailbox by default. Passing a mailbox UPN
 * targets a shared/delegated mailbox that the signed-in user can access.
 */
export function owaMailboxRoot(mailbox?: string): string {
  if (!mailbox) return `${OWA_REST_V2}/me`;

  const normalized = mailbox.trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalized)) {
    throw new Error('mailbox must be a valid email address.');
  }

  return `${OWA_REST_V2}/users/${encodeURIComponent(normalized)}`;
}

export async function owaGet<T>(path: string, params?: Record<string, string>, mailbox?: string): Promise<T> {
  const token = await getOwaToken();
  if (!token) throw new Error('Not authenticated. Run outlook_login first.');

  let url = `${owaMailboxRoot(mailbox)}${path}`;
  if (params) {
    const qs = new URLSearchParams(params).toString();
    url += (url.includes('?') ? '&' : '?') + qs;
  }

  const res = await fetchWithRetry(url, {
    method: 'GET',
    headers: getBearerHeaders(token, OWA_BASE),
  });
  return parseResponse<T>(res);
}

export async function owaPost<T>(path: string, body: unknown, mailbox?: string): Promise<T> {
  const token = await getOwaToken();
  if (!token) throw new Error('Not authenticated. Run outlook_login first.');

  const res = await fetchWithRetry(`${owaMailboxRoot(mailbox)}${path}`, {
    method: 'POST',
    headers: getBearerHeaders(token, OWA_BASE),
    body: JSON.stringify(body),
  });
  return parseResponse<T>(res);
}

export async function owaPatch<T>(path: string, body: unknown, mailbox?: string): Promise<T> {
  const token = await getOwaToken();
  if (!token) throw new Error('Not authenticated. Run outlook_login first.');

  const res = await fetchWithRetry(`${owaMailboxRoot(mailbox)}${path}`, {
    method: 'PATCH',
    headers: getBearerHeaders(token, OWA_BASE),
    body: JSON.stringify(body),
  });
  return parseResponse<T>(res);
}

export async function owaDelete(path: string, mailbox?: string): Promise<void> {
  const token = await getOwaToken();
  if (!token) throw new Error('Not authenticated. Run outlook_login first.');

  const res = await fetchWithRetry(`${owaMailboxRoot(mailbox)}${path}`, {
    method: 'DELETE',
    headers: getBearerHeaders(token, OWA_BASE),
  });
  if (!res.ok && res.status !== 204) {
    const text = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status}: ${text.slice(0, 200)}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Graph API client (fallback for some endpoints)
// ─────────────────────────────────────────────────────────────────────────────

export async function graphGet<T>(path: string, params?: Record<string, string>): Promise<T> {
  const token = await getGraphToken();
  if (!token) throw new Error('Graph token unavailable. Run outlook_login first.');

  let url = `${GRAPH_BASE}/me${path}`;
  if (params) {
    const qs = new URLSearchParams(params).toString();
    url += (url.includes('?') ? '&' : '?') + qs;
  }

  const res = await fetchWithRetry(url, {
    method: 'GET',
    headers: getBearerHeaders(token),
  });
  return parseResponse<T>(res);
}

export async function graphPost<T>(path: string, body: unknown): Promise<T> {
  const token = await getGraphToken();
  if (!token) throw new Error('Graph token unavailable. Run outlook_login first.');

  const res = await fetchWithRetry(`${GRAPH_BASE}/me${path}`, {
    method: 'POST',
    headers: getBearerHeaders(token),
    body: JSON.stringify(body),
  });
  return parseResponse<T>(res);
}

/**
 * Graph GET against an arbitrary path (not forced under /me), e.g.
 * '/users/{email}', '/users/{email}/manager', '/users/{email}/directReports'.
 * Callers include the full path after the version segment.
 */
export async function graphGetPath<T>(path: string, params?: Record<string, string>): Promise<T> {
  const token = await getGraphToken();
  if (!token) throw new Error('Graph token unavailable. Run outlook_login first.');

  let url = `${GRAPH_BASE}${path}`;
  if (params) {
    const qs = new URLSearchParams(params).toString();
    url += (url.includes('?') ? '&' : '?') + qs;
  }

  const res = await fetchWithRetry(url, { method: 'GET', headers: getBearerHeaders(token) });
  return parseResponse<T>(res);
}

/** Graph GET that returns raw bytes (e.g. a profile photo at /photo/$value). */
export async function graphGetBinary(path: string): Promise<{ contentType: string; bytes: Buffer }> {
  const token = await getGraphToken();
  if (!token) throw new Error('Graph token unavailable. Run outlook_login first.');

  const res = await fetchWithRetry(`${GRAPH_BASE}${path}`, { method: 'GET', headers: getBearerHeaders(token) });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status} ${res.statusText}: ${text.slice(0, 200)}`);
  }
  const contentType = res.headers.get('content-type') ?? 'application/octet-stream';
  const bytes = Buffer.from(await res.arrayBuffer());
  return { contentType, bytes };
}
