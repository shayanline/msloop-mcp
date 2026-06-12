/**
 * Loop page content export.
 *
 * Page content lives in a Fluid document in SharePoint. There is no plain
 * content-read API, but SharePoint can render the document to HTML on demand
 * via the `?format=html` query parameter. We fetch that HTML and convert it to
 * Markdown for LLM consumption.
 */

import { sharePointGetText } from './client.js';
import { itemIdFromPageId, hostFromSiteUrl, htmlToMarkdown } from '../utils/parsers.js';
import type { LoopPage, SpoCoordinates } from '../types/loop.js';

export interface PageContent {
  html: string;
  markdown: string;
  coordinates: SpoCoordinates;
}

function contentUrl({ host, driveId, itemId }: SpoCoordinates): string {
  return `https://${host}/_api/v2.0/drives/${driveId}/items/${itemId}/content?format=html&ump=1`;
}

/**
 * Resolve the SharePoint coordinates for a page, optionally using workspace
 * fallback coordinates (from the workspace pod_id) when the page metadata is
 * incomplete.
 */
export function resolvePageCoordinates(
  page: LoopPage,
  fallback?: SpoCoordinates | null,
): SpoCoordinates | null {
  const host = hostFromSiteUrl(page.sharepoint_info?.site_url) ?? fallback?.host;
  const driveId = page.onedrive_info?.drive_id ?? fallback?.driveId;
  const itemId = itemIdFromPageId(page.id) || fallback?.itemId;
  if (!host || !driveId || !itemId) return null;
  return { host, driveId, itemId };
}

/** Fetch a page's content as HTML and Markdown. */
export async function getPageContent(
  page: LoopPage,
  fallback?: SpoCoordinates | null,
): Promise<PageContent> {
  const coordinates = resolvePageCoordinates(page, fallback);
  if (!coordinates) {
    throw new Error(
      `Cannot locate SharePoint storage for page "${page.title ?? page.id}". ` +
      'The page metadata is missing drive/site info and no workspace fallback was available.',
    );
  }
  const html = await sharePointGetText(contentUrl(coordinates));
  return { html, markdown: htmlToMarkdown(html), coordinates };
}
