/**
 * HTTP utility helpers.
 */

import { randomUUID } from 'node:crypto';
import { LOOP_ORIGIN, LOOP_USER_AGENT } from '../constants.js';

/** Standard JSON bearer headers for the Substrate and Graph APIs. */
export function getBearerHeaders(token: string, origin = LOOP_ORIGIN): Record<string, string> {
  return {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
    Accept: 'application/json',
    Origin: origin,
    'User-Agent': LOOP_USER_AGENT,
  };
}

/**
 * Parse a fetch Response, throwing on non-2xx. JSON bodies are parsed; empty or
 * non-JSON bodies (e.g. 202/204) resolve to undefined.
 */
export async function parseResponse<T>(res: Response): Promise<T> {
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status} ${res.statusText}: ${text.slice(0, 300)}`);
  }
  const ct = res.headers.get('content-type') ?? '';
  if (ct.includes('application/json')) {
    return res.json() as Promise<T>;
  }
  return undefined as T;
}

/** Sleep for ms milliseconds. */
export const sleep = (ms: number) => new Promise<void>(r => setTimeout(r, ms));

/** Retry a fetch operation with exponential backoff on 429/503/504. */
export async function fetchWithRetry(
  url: string,
  init: RequestInit,
  maxAttempts = 3,
): Promise<Response> {
  let lastError: Error | null = null;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    const res = await fetch(url, init);
    if (res.status === 429 || res.status === 503 || res.status === 504) {
      const retryAfter = parseInt(res.headers.get('Retry-After') ?? '1', 10);
      const delay = (retryAfter || Math.pow(2, attempt)) * 1000 + Math.random() * 500;
      await sleep(delay);
      lastError = new Error(`HTTP ${res.status}`);
      continue;
    }
    return res;
  }
  throw lastError ?? new Error('Max retries exceeded');
}

// ─────────────────────────────────────────────────────────────────────────────
// SharePoint multipart GET
// ─────────────────────────────────────────────────────────────────────────────

/**
 * SharePoint Fluid endpoints don't accept a plain Authorization header from a
 * browser origin. The Loop web client smuggles the credentials inside a
 * multipart/form-data POST, with `X-HTTP-Method-Override: GET` telling
 * SharePoint to treat it as a GET. This builds that request body.
 */
export function buildSharePointGetBody(token: string, boundary: string): string {
  return [
    `--${boundary}`,
    `Authorization: Bearer ${token}`,
    'X-HTTP-Method-Override: GET',
    '_post: 1',
    '',
    `--${boundary}--`,
    '',
  ].join('\r\n');
}

/**
 * Perform a SharePoint "GET via multipart POST" request and return the raw
 * Response (caller decides whether to read .text() or .json()).
 */
export async function sharePointGet(url: string, token: string): Promise<Response> {
  const boundary = randomUUID();
  return fetchWithRetry(url, {
    method: 'POST',
    headers: {
      'Content-Type': `multipart/form-data; boundary=${boundary}`,
      Origin: LOOP_ORIGIN,
      Referer: `${LOOP_ORIGIN}/`,
      'User-Agent': LOOP_USER_AGENT,
    },
    body: buildSharePointGetBody(token, boundary),
  });
}
