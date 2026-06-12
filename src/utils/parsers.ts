/**
 * Pure, testable parsing helpers for Loop data.
 */

import TurndownService from 'turndown';
import type { SpoCoordinates } from '../types/loop.js';

/**
 * Decode a workspace `mfs_info.pod_id` into SharePoint coordinates.
 *
 * The pod_id is a base64 string whose decoded form is a pipe-delimited list
 * ending in `…|{host}|{driveId}|{itemId}`. We take the last three segments.
 * Returns null if the value can't be decoded into three trailing segments.
 */
export function decodePodId(podId: string | undefined): SpoCoordinates | null {
  if (!podId) return null;
  let decoded: string;
  try {
    decoded = Buffer.from(podId, 'base64').toString('utf8');
  } catch {
    return null;
  }
  const parts = decoded.split('|').filter(Boolean);
  if (parts.length < 3) return null;
  const [host, driveId, itemId] = parts.slice(-3);
  if (!host || !driveId || !itemId) return null;
  return { host, driveId, itemId };
}

/**
 * Extract the SharePoint item ID from a Loop page ID. Loop page IDs end with
 * the SPO item ID after the final underscore, e.g. "…_01ABCXYZ" → "01ABCXYZ".
 * Falls back to the whole id when there is no underscore.
 */
export function itemIdFromPageId(pageId: string): string {
  const idx = pageId.lastIndexOf('_');
  return idx >= 0 ? pageId.slice(idx + 1) : pageId;
}

/** Derive the SharePoint host (no scheme) from a site URL like https://contoso.sharepoint.com/sites/x. */
export function hostFromSiteUrl(siteUrl: string | undefined): string | null {
  if (!siteUrl) return null;
  try {
    return new URL(siteUrl).host;
  } catch {
    // siteUrl may already be a bare host
    const m = siteUrl.match(/[^/]+\.sharepoint\.com/i);
    return m ? m[0] : null;
  }
}

/** Slugify a title into a filesystem/url-safe segment. */
export function slugify(title: string): string {
  return title
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    || 'untitled';
}

let turndown: TurndownService | null = null;

/** Convert a Loop page's exported HTML into Markdown. */
export function htmlToMarkdown(html: string): string {
  if (!turndown) {
    turndown = new TurndownService({
      headingStyle: 'atx',
      codeBlockStyle: 'fenced',
      bulletListMarker: '-',
    });
  }
  return turndown.turndown(html).trim();
}
