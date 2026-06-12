/**
 * Base API clients for the three services Loop relies on:
 *   - Substrate Loop API  (workspace + page metadata)   — JSON bearer
 *   - Microsoft Graph      (file search + metadata)       — JSON bearer
 *   - SharePoint            (Fluid snapshots + page HTML)  — multipart "GET via POST"
 */

import { getSubstrateToken, getSharePointToken, getGraphToken } from '../auth/index.js';
import { getBearerHeaders, parseResponse, fetchWithRetry, sharePointGet } from '../utils/http.js';

// ─────────────────────────────────────────────────────────────────────────────
// Substrate Loop API
// ─────────────────────────────────────────────────────────────────────────────

export async function substrateGet<T>(url: string): Promise<T> {
  const token = await getSubstrateToken();
  if (!token) throw new Error('Not authenticated for Substrate. Run loop_login first.');
  const res = await fetchWithRetry(url, { method: 'GET', headers: getBearerHeaders(token) });
  return parseResponse<T>(res);
}

export async function substratePost<T>(url: string, body: unknown): Promise<T> {
  const token = await getSubstrateToken();
  if (!token) throw new Error('Not authenticated for Substrate. Run loop_login first.');
  const res = await fetchWithRetry(url, {
    method: 'POST',
    headers: getBearerHeaders(token),
    body: JSON.stringify(body),
  });
  return parseResponse<T>(res);
}

// ─────────────────────────────────────────────────────────────────────────────
// Microsoft Graph
// ─────────────────────────────────────────────────────────────────────────────

export async function graphGet<T>(url: string): Promise<T> {
  const token = await getGraphToken();
  if (!token) throw new Error('Graph token unavailable. Run loop_login first.');
  const res = await fetchWithRetry(url, { method: 'GET', headers: getBearerHeaders(token) });
  return parseResponse<T>(res);
}

export async function graphPost<T>(url: string, body: unknown): Promise<T> {
  const token = await getGraphToken();
  if (!token) throw new Error('Graph token unavailable. Run loop_login first.');
  const res = await fetchWithRetry(url, {
    method: 'POST',
    headers: getBearerHeaders(token),
    body: JSON.stringify(body),
  });
  return parseResponse<T>(res);
}

// ─────────────────────────────────────────────────────────────────────────────
// SharePoint (Fluid snapshots + page content)
// ─────────────────────────────────────────────────────────────────────────────

/** Fetch a SharePoint resource as text via the multipart "GET via POST" convention. */
export async function sharePointGetText(url: string): Promise<string> {
  const token = await getSharePointToken();
  if (!token) throw new Error('SharePoint token unavailable. Run loop_login first.');
  const res = await sharePointGet(url, token);
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`SharePoint HTTP ${res.status} ${res.statusText}: ${text.slice(0, 300)}`);
  }
  return res.text();
}
