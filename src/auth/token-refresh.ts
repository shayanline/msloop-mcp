/**
 * HTTP-based token refresh.
 *
 * Uses the Loop first-party client ID (public SPA — no client secret needed) to
 * exchange a cached refresh token for new access tokens via the standard OAuth2
 * token endpoint, one call per resource (Substrate, SharePoint, Graph).
 *
 * Key detail (from msteams-mcp): the Origin header is REQUIRED for SPA client
 * IDs. Azure AD validates that refresh-token grants from SPA clients include a
 * cross-origin Origin header matching a registered redirect URI; without it
 * Azure AD returns AADSTS9002327.
 */

import { logger } from '../utils/logger.js';
import {
  LOOP_CLIENT_ID,
  LOOP_ORIGIN,
  SUBSTRATE_SCOPE,
  GRAPH_SCOPE,
  TOKEN_ENDPOINT_TEMPLATE,
} from '../constants.js';
import { readTokenCache, writeTokenCache, type TokenCache } from './session-store.js';

const REFRESH_TIMEOUT_MS = 10_000;

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  scope?: string;
}

const inProgress: Record<string, boolean> = {};

async function callTokenEndpoint(
  tenantId: string,
  refreshToken: string,
  scope: string,
): Promise<TokenResponse | null> {
  const url = TOKEN_ENDPOINT_TEMPLATE.replace('{tenant}', tenantId);

  const body = new URLSearchParams({
    client_id: LOOP_CLIENT_ID,
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope,
  });

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), REFRESH_TIMEOUT_MS);

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        // Required for SPA public clients — Azure AD returns AADSTS9002327 without this.
        'Origin': LOOP_ORIGIN,
      },
      body: body.toString(),
      signal: controller.signal,
    });

    if (!res.ok) {
      const text = await res.text().catch(() => '');
      logger.warn(`Token refresh HTTP ${res.status} (scope ${scope})`, text.slice(0, 200));
      return null;
    }

    return await res.json() as TokenResponse;
  } catch (err) {
    if (err instanceof Error && err.name === 'AbortError') {
      logger.warn('Token refresh request timed out');
    } else {
      logger.warn('Token refresh network error', err instanceof Error ? err.message : String(err));
    }
    return null;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Refresh an access token for the given resource scope and persist it into the
 * cache via `apply`. Returns the new access token on success, null on failure.
 */
async function refreshResource(
  guardKey: string,
  scope: string,
  apply: (cache: TokenCache, token: string, expiry: number) => TokenCache,
): Promise<string | null> {
  if (inProgress[guardKey]) {
    logger.debug(`Refresh already in progress for ${guardKey}`);
    return null;
  }

  const cache = readTokenCache();
  if (!cache?.refreshToken || !cache.tenantId) {
    logger.debug('No refresh token or tenant ID cached');
    return null;
  }

  inProgress[guardKey] = true;
  try {
    const response = await callTokenEndpoint(cache.tenantId, cache.refreshToken, scope);
    if (!response) return null;

    const expiry = Date.now() + response.expires_in * 1000;
    const updated = apply(
      { ...cache, refreshToken: response.refresh_token ?? cache.refreshToken, extractedAt: Date.now() },
      response.access_token,
      expiry,
    );
    writeTokenCache(updated);
    logger.info(`Refreshed ${guardKey} token successfully`);
    return response.access_token;
  } finally {
    inProgress[guardKey] = false;
  }
}

export function refreshSubstrateToken(): Promise<string | null> {
  return refreshResource('substrate', SUBSTRATE_SCOPE, (cache, token, expiry) => ({
    ...cache,
    substrateToken: token,
    substrateTokenExpiry: expiry,
  }));
}

export function refreshGraphToken(): Promise<string | null> {
  return refreshResource('graph', GRAPH_SCOPE, (cache, token, expiry) => ({
    ...cache,
    graphToken: token,
    graphTokenExpiry: expiry,
  }));
}

/**
 * Refresh the SharePoint token. Requires a known SharePoint resource host
 * (e.g. https://contoso.sharepoint.com) discovered during login; without it
 * we cannot form the resource scope.
 */
export function refreshSharePointToken(): Promise<string | null> {
  const cache = readTokenCache();
  if (!cache?.sharePointResource) {
    logger.debug('No SharePoint resource cached — cannot refresh SharePoint token');
    return Promise.resolve(null);
  }
  const scope = `${cache.sharePointResource}/.default`;
  return refreshResource('sharepoint', scope, (c, token, expiry) => ({
    ...c,
    sharePointToken: token,
    sharePointTokenExpiry: expiry,
  }));
}
