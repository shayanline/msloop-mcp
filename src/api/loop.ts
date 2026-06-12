/**
 * Substrate Loop API — workspace and page discovery, and workspace creation.
 *
 * No single endpoint returns everything, so workspace/page discovery queries
 * three endpoints and merges the results, deduplicating by `id`:
 *   - /workspaces  — canonical list (only reliable source of "My workspace")
 *   - /recent      — recently active workspaces + pages
 *   - /deltasync   — full component graph (bulk of page metadata)
 */

import { LOOP_API_BASE, SPEEDWAY_BASE, SUBSTRATE_BASE } from '../constants.js';
import { logger } from '../utils/logger.js';
import { substrateGet, substratePost } from './client.js';
import type { LoopData, LoopWorkspace, LoopPage } from '../types/loop.js';

/** Cap on pagination follow-ups, to avoid runaway loops. */
const MAX_PAGES = 20;

const ENDPOINTS = [
  `${LOOP_API_BASE}/workspaces?rs=en-us`,
  `${LOOP_API_BASE}/recent?top=30&settings=true&rs=en-us`,
  `${LOOP_API_BASE}/deltasync?loopComponents=true&rs=en-us`,
];

/** Resolve a next_page_link (which may be relative) against the Substrate base. */
function resolveNextLink(link: string): string {
  if (/^https?:\/\//i.test(link)) return link;
  return link.startsWith('/') ? `${SUBSTRATE_BASE}${link}` : `${LOOP_API_BASE}/${link}`;
}

async function fetchAllPages(startUrl: string): Promise<LoopData[]> {
  const results: LoopData[] = [];
  let url: string | undefined = startUrl;
  let count = 0;
  while (url && count < MAX_PAGES) {
    const data: LoopData = await substrateGet<LoopData>(url);
    results.push(data);
    url = data.next_page_link ? resolveNextLink(data.next_page_link) : undefined;
    count++;
  }
  return results;
}

interface DiscoverResult {
  workspaces: LoopWorkspace[];
  pages: LoopPage[];
}

/** Query all three discovery endpoints and merge, deduplicating by id. */
export async function discover(): Promise<DiscoverResult> {
  const workspaces = new Map<string, LoopWorkspace>();
  const pages = new Map<string, LoopPage>();

  const settled = await Promise.allSettled(ENDPOINTS.map(fetchAllPages));

  for (const result of settled) {
    if (result.status === 'rejected') {
      logger.debug('A Loop discovery endpoint failed', String(result.reason));
      continue;
    }
    for (const data of result.value) {
      for (const ws of data.workspaces ?? []) {
        if (ws.id && !workspaces.has(ws.id)) workspaces.set(ws.id, ws);
      }
      for (const pg of data.pages ?? []) {
        if (pg.id && !pages.has(pg.id)) pages.set(pg.id, pg);
      }
    }
  }

  if (workspaces.size === 0 && pages.size === 0) {
    throw new Error('No Loop data returned. The session may be invalid — try loop_login.');
  }

  return { workspaces: [...workspaces.values()], pages: [...pages.values()] };
}

/** List all workspaces the user can see. */
export async function listWorkspaces(): Promise<LoopWorkspace[]> {
  return (await discover()).workspaces;
}

/** List the (non-deleted) pages belonging to a given workspace. */
export async function listPages(workspaceId: string, includeDeleted = false): Promise<LoopPage[]> {
  const { pages } = await discover();
  return pages.filter(
    p => p.workspace_id === workspaceId && (includeDeleted || !p.is_deleted),
  );
}

/** Look up a single page by its id across all workspaces. */
export async function findPage(pageId: string): Promise<LoopPage | null> {
  const { pages } = await discover();
  return pages.find(p => p.id === pageId) ?? null;
}

// ─────────────────────────────────────────────────────────────────────────────
// Workspace creation (experimental)
// ─────────────────────────────────────────────────────────────────────────────

export interface CreatedWorkspace {
  id?: string;
  displayName?: string;
  [k: string]: unknown;
}

/**
 * Create a new shared Loop workspace via the Substrate Speedway API.
 *
 * EXPERIMENTAL: workspace creation is only partially documented and Loop may
 * require follow-up provisioning before the workspace is fully usable in the
 * web app. The call returns whatever the API responds with.
 */
export async function createWorkspace(displayName: string): Promise<CreatedWorkspace> {
  const body = {
    displayName,
    isPersonal: false,
    enabledWorkloads: ['SharePoint'],
    groupType: 'Workspace',
  };
  return substratePost<CreatedWorkspace>(`${SPEEDWAY_BASE}/workspaceGroups`, body);
}
