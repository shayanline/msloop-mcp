/**
 * Extract MSAL tokens from the Loop web app's storage.
 *
 * The Loop SPA (loop.cloud.microsoft) uses MSAL and stores access tokens,
 * refresh tokens, and ID tokens under keys like:
 *   <accountId>-login.windows.net-accesstoken-<clientId>-<tenantId>-<scopes>
 *   <accountId>-login.windows.net-refreshtoken-<clientId>----
 *
 * Depending on the MSAL `cacheLocation`, these live in localStorage or
 * sessionStorage, so the browser-login flow gathers entries from BOTH stores
 * and passes the merged list here. Tokens are then selected by their audience:
 *   - Substrate  → target/aud contains "substrate.office.com"
 *   - SharePoint → target/aud contains ".sharepoint.com"
 *   - Graph      → target/aud contains "graph.microsoft.com"
 *
 * This mirrors how msteams-mcp and msoutlook-mcp extract tokens, adapted for Loop.
 */

import { logger } from '../utils/logger.js';
import { LOOP_CLIENT_ID } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface ExtractedTokens {
  substrateToken?: string;
  substrateTokenExpiry?: Date;
  sharePointToken?: string;
  sharePointTokenExpiry?: Date;
  sharePointResource?: string;
  graphToken?: string;
  graphTokenExpiry?: Date;
  refreshToken: string;
  tenantId?: string;
  upn?: string;
}

interface MsalEntry {
  secret: string;
  credentialType?: string;
  target?: string;
  realm?: string;
  homeAccountId?: string;
  clientId?: string;
  environment?: string;
}

interface JwtPayload {
  exp?: number;
  aud?: string;
  upn?: string;
  preferred_username?: string;
  tid?: string;
  appid?: string;
}

export interface StorageEntry { name: string; value: string }

// ─────────────────────────────────────────────────────────────────────────────
// JWT utilities
// ─────────────────────────────────────────────────────────────────────────────

export function decodeJwtPayload(token: string): JwtPayload | null {
  try {
    const parts = token.split('.');
    if (parts.length < 2) return null;
    const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
    return JSON.parse(Buffer.from(padded, 'base64').toString('utf8')) as JwtPayload;
  } catch {
    return null;
  }
}

function getJwtExpiry(token: string): Date | null {
  const payload = decodeJwtPayload(token);
  if (!payload?.exp || typeof payload.exp !== 'number') return null;
  return new Date(payload.exp * 1000);
}

function isJwt(value: string): boolean {
  return typeof value === 'string' && value.startsWith('ey');
}

/** Pull the resource host (e.g. https://contoso.sharepoint.com) out of a token's audience. */
export function resourceFromToken(token: string): string | undefined {
  const aud = decodeJwtPayload(token)?.aud;
  if (!aud) return undefined;
  // aud may be a bare GUID, a full URL, or a URL with a trailing path/guid.
  const match = aud.match(/^https?:\/\/[^/]+/);
  return match ? match[0] : undefined;
}

// ─────────────────────────────────────────────────────────────────────────────
// Extraction
// ─────────────────────────────────────────────────────────────────────────────

interface Candidate { token: string; expiry: Date }

function betterOf(current: Candidate | null, next: Candidate): Candidate {
  return !current || next.expiry > current.expiry ? next : current;
}

/**
 * Extract all relevant MSAL tokens from a merged list of storage entries
 * (localStorage + sessionStorage). Tokens are matched by audience, not by a
 * single hard-coded scope, so this is resilient to tenant differences.
 */
export function extractTokensFromEntries(entries: StorageEntry[]): ExtractedTokens | null {
  let bestSubstrate: Candidate | null = null;
  let bestSharePoint: Candidate | null = null;
  let bestGraph: Candidate | null = null;
  let sharePointResource: string | undefined;
  let refreshToken: string | null = null;
  let tenantId: string | undefined;
  let upn: string | undefined;

  for (const item of entries) {
    const value = item.value;
    if (typeof value !== 'string') continue;

    let entry: MsalEntry;
    try {
      entry = JSON.parse(value) as MsalEntry;
    } catch {
      continue;
    }
    if (!entry.secret) continue;

    const key = item.name.toLowerCase();

    // ── Refresh token ──────────────────────────────────────────────────────
    if (key.includes('refreshtoken')) {
      // Prefer the Loop client's refresh token; fall back to any if none matched yet.
      if (entry.clientId === LOOP_CLIENT_ID || !refreshToken) {
        refreshToken = entry.secret;
      }
      continue;
    }

    // ── Access tokens ──────────────────────────────────────────────────────
    if (!key.includes('accesstoken')) continue;
    if (!isJwt(entry.secret)) continue;

    const expiry = getJwtExpiry(entry.secret);
    if (!expiry || expiry.getTime() <= Date.now()) continue; // skip expired

    const payload = decodeJwtPayload(entry.secret);
    if (!upn) upn = payload?.upn ?? payload?.preferred_username;
    if (!tenantId) tenantId = payload?.tid;

    // Match by target (the MSAL scope) or by the token audience.
    const haystack = `${entry.target ?? ''} ${payload?.aud ?? ''}`.toLowerCase();

    if (haystack.includes('substrate.office.com')) {
      bestSubstrate = betterOf(bestSubstrate, { token: entry.secret, expiry });
    } else if (haystack.includes('.sharepoint.com')) {
      const cand = { token: entry.secret, expiry };
      if (betterOf(bestSharePoint, cand) === cand) {
        bestSharePoint = cand;
        sharePointResource = resourceFromToken(entry.secret) ?? sharePointResource;
      }
    } else if (haystack.includes('graph.microsoft.com')) {
      bestGraph = betterOf(bestGraph, { token: entry.secret, expiry });
    }
  }

  if (!refreshToken) {
    logger.debug('Token extraction failed: no refresh token found');
    return null;
  }
  if (!bestSubstrate && !bestSharePoint && !bestGraph) {
    logger.debug('Token extraction failed: no Loop-relevant access tokens found');
    return null;
  }

  return {
    substrateToken: bestSubstrate?.token,
    substrateTokenExpiry: bestSubstrate?.expiry,
    sharePointToken: bestSharePoint?.token,
    sharePointTokenExpiry: bestSharePoint?.expiry,
    sharePointResource,
    graphToken: bestGraph?.token,
    graphTokenExpiry: bestGraph?.expiry,
    refreshToken,
    tenantId,
    upn,
  };
}
