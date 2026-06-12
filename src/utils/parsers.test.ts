import { describe, it, expect } from 'vitest';
import {
  decodePodId,
  itemIdFromPageId,
  hostFromSiteUrl,
  slugify,
  htmlToMarkdown,
} from './parsers.js';

const b64 = (s: string) => Buffer.from(s, 'utf8').toString('base64');

describe('decodePodId', () => {
  it('takes the last three pipe-delimited segments', () => {
    const podId = b64('1|something|contoso.sharepoint.com|b!driveId123|01ITEMABC');
    expect(decodePodId(podId)).toEqual({
      host: 'contoso.sharepoint.com',
      driveId: 'b!driveId123',
      itemId: '01ITEMABC',
    });
  });

  it('returns null for undefined or too-few segments', () => {
    expect(decodePodId(undefined)).toBeNull();
    expect(decodePodId(b64('host|drive'))).toBeNull();
  });
});

describe('itemIdFromPageId', () => {
  it('returns the segment after the last underscore', () => {
    expect(itemIdFromPageId('abc_def_01ITEM')).toBe('01ITEM');
  });
  it('returns the whole id when there is no underscore', () => {
    expect(itemIdFromPageId('plainid')).toBe('plainid');
  });
});

describe('hostFromSiteUrl', () => {
  it('extracts host from a full URL', () => {
    expect(hostFromSiteUrl('https://contoso.sharepoint.com/sites/x')).toBe('contoso.sharepoint.com');
  });
  it('extracts a bare host', () => {
    expect(hostFromSiteUrl('contoso.sharepoint.com')).toBe('contoso.sharepoint.com');
  });
  it('returns null for undefined', () => {
    expect(hostFromSiteUrl(undefined)).toBeNull();
  });
});

describe('slugify', () => {
  it('lowercases and dashes', () => {
    expect(slugify('My Cool Page!')).toBe('my-cool-page');
  });
  it('falls back to untitled', () => {
    expect(slugify('!!!')).toBe('untitled');
  });
});

describe('htmlToMarkdown', () => {
  it('converts basic HTML to markdown', () => {
    const md = htmlToMarkdown('<h1>Title</h1><p>Hello <strong>world</strong></p>');
    expect(md).toContain('# Title');
    expect(md).toContain('**world**');
  });
});
