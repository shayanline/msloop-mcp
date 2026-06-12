import { describe, it, expect } from 'vitest';
import { buildSharePointGetBody, getBearerHeaders } from './http.js';

describe('buildSharePointGetBody', () => {
  it('builds a multipart body that smuggles the bearer token and overrides the method to GET', () => {
    const body = buildSharePointGetBody('tok123', 'BOUNDARY');
    expect(body).toContain('--BOUNDARY');
    expect(body).toContain('Authorization: Bearer tok123');
    expect(body).toContain('X-HTTP-Method-Override: GET');
    expect(body.trimEnd().endsWith('--BOUNDARY--')).toBe(true);
  });
});

describe('getBearerHeaders', () => {
  it('sets the auth header and a loop origin by default', () => {
    const h = getBearerHeaders('abc');
    expect(h.Authorization).toBe('Bearer abc');
    expect(h.Origin).toContain('loop.cloud.microsoft');
  });
});
