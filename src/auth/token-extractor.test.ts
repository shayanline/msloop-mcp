import { describe, it, expect } from 'vitest';
import { extractTokensFromEntries, resourceFromToken, decodeJwtPayload, type StorageEntry } from './token-extractor.js';
import { LOOP_CLIENT_ID } from '../constants.js';

/** Build a fake JWT with the given payload (signature is irrelevant here). */
function jwt(payload: Record<string, unknown>): string {
  const enc = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${enc({ alg: 'none' })}.${enc(payload)}.sig`;
}

const future = Math.floor(Date.now() / 1000) + 3600;
const past = Math.floor(Date.now() / 1000) - 3600;

function accessEntry(name: string, target: string, payload: Record<string, unknown>): StorageEntry {
  return {
    name,
    value: JSON.stringify({ credentialType: 'AccessToken', target, secret: jwt(payload) }),
  };
}

describe('decodeJwtPayload', () => {
  it('decodes a base64url payload', () => {
    expect(decodeJwtPayload(jwt({ tid: 'abc' }))?.tid).toBe('abc');
  });
  it('returns null for junk', () => {
    expect(decodeJwtPayload('not-a-jwt')).toBeNull();
  });
});

describe('resourceFromToken', () => {
  it('extracts the origin from a URL audience', () => {
    expect(resourceFromToken(jwt({ aud: 'https://contoso.sharepoint.com/foo' }))).toBe('https://contoso.sharepoint.com');
  });
  it('returns undefined for a bare GUID audience', () => {
    expect(resourceFromToken(jwt({ aud: '00000003-0000-0000-c000-000000000000' }))).toBeUndefined();
  });
});

describe('extractTokensFromEntries', () => {
  it('selects tokens by audience and finds the refresh token', () => {
    const entries: StorageEntry[] = [
      accessEntry('a-accesstoken-substrate', 'https://substrate.office.com/.default', {
        exp: future, aud: 'https://substrate.office.com', tid: 'tenant-1', upn: 'me@contoso.com',
      }),
      accessEntry('a-accesstoken-spo', 'https://contoso.sharepoint.com/.default', {
        exp: future, aud: 'https://contoso.sharepoint.com',
      }),
      accessEntry('a-accesstoken-graph', 'https://graph.microsoft.com/.default', {
        exp: future, aud: 'https://graph.microsoft.com',
      }),
      { name: 'a-refreshtoken-loop', value: JSON.stringify({ credentialType: 'RefreshToken', clientId: LOOP_CLIENT_ID, secret: 'rt-value' }) },
    ];

    const result = extractTokensFromEntries(entries);
    expect(result).not.toBeNull();
    expect(result!.substrateToken).toBeTruthy();
    expect(result!.sharePointToken).toBeTruthy();
    expect(result!.sharePointResource).toBe('https://contoso.sharepoint.com');
    expect(result!.graphToken).toBeTruthy();
    expect(result!.refreshToken).toBe('rt-value');
    expect(result!.tenantId).toBe('tenant-1');
    expect(result!.upn).toBe('me@contoso.com');
  });

  it('skips expired access tokens', () => {
    const entries: StorageEntry[] = [
      accessEntry('a-accesstoken-substrate', 'https://substrate.office.com/.default', { exp: past, aud: 'https://substrate.office.com' }),
      { name: 'a-refreshtoken', value: JSON.stringify({ credentialType: 'RefreshToken', clientId: LOOP_CLIENT_ID, secret: 'rt' }) },
    ];
    // No live access token but a refresh token exists → null (nothing usable yet).
    expect(extractTokensFromEntries(entries)).toBeNull();
  });

  it('returns null when there is no refresh token', () => {
    const entries: StorageEntry[] = [
      accessEntry('a-accesstoken-substrate', 'https://substrate.office.com/.default', { exp: future, aud: 'https://substrate.office.com' }),
    ];
    expect(extractTokensFromEntries(entries)).toBeNull();
  });

  it('prefers the highest-expiry token for a resource', () => {
    const soon = Math.floor(Date.now() / 1000) + 600;
    const entries: StorageEntry[] = [
      accessEntry('accesstoken-t1', 'https://substrate.office.com/.default', { exp: soon, aud: 'https://substrate.office.com' }),
      accessEntry('accesstoken-t2', 'https://substrate.office.com/.default', { exp: future, aud: 'https://substrate.office.com' }),
      { name: 'a-refreshtoken', value: JSON.stringify({ credentialType: 'RefreshToken', clientId: LOOP_CLIENT_ID, secret: 'rt' }) },
    ];
    const result = extractTokensFromEntries(entries);
    expect(result!.substrateTokenExpiry!.getTime()).toBe(future * 1000);
  });
});
