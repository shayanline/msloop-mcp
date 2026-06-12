/**
 * Page MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { discover } from '../api/loop.js';
import { getPageContent } from '../api/pages.js';
import { decodePodId } from '../utils/parsers.js';
import type { LoopPage, LoopWorkspace } from '../types/loop.js';

function summarisePage(p: LoopPage) {
  return {
    id: p.id,
    title: p.title ?? '(untitled)',
    type: p.type,
    workspaceId: p.workspace_id,
    isDeleted: p.is_deleted ?? false,
  };
}

/** Fallback SharePoint coordinates for a page, from its workspace pod_id. */
function workspaceFallback(workspaces: LoopWorkspace[], workspaceId: string | undefined) {
  if (!workspaceId) return null;
  const ws = workspaces.find(w => w.id === workspaceId);
  return decodePodId(ws?.mfs_info?.pod_id);
}

export function registerPageTools(server: McpServer): void {
  // ── loop_list_pages ──────────────────────────────────────────────────────
  server.tool(
    'loop_list_pages',
    'List the pages in a Loop workspace. Pass the workspace id from loop_list_workspaces.',
    {
      workspace_id: z.string().min(1).describe('The workspace id to list pages for.'),
      include_deleted: z.boolean().optional().describe('Include deleted pages. Default: false.'),
    },
    async ({ workspace_id, include_deleted }) => {
      const { pages } = await discover();
      const filtered = pages.filter(
        p => p.workspace_id === workspace_id && (include_deleted || !p.is_deleted),
      );
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({ count: filtered.length, pages: filtered.map(summarisePage) }, null, 2),
        }],
      };
    },
  );

  // ── loop_get_page ────────────────────────────────────────────────────────
  server.tool(
    'loop_get_page',
    'Read the content of a Loop page as Markdown. Pass the page id from loop_list_pages. Loop pages are rich Fluid documents exported to HTML, so some interactive components may render approximately.',
    {
      page_id: z.string().min(1).describe('The page id to read.'),
      format: z.enum(['markdown', 'html']).optional().describe('Output format. Default: markdown.'),
    },
    async ({ page_id, format }) => {
      const { pages, workspaces } = await discover();
      const page = pages.find(p => p.id === page_id);
      if (!page) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({ success: false, message: `No page found with id ${page_id}.` }, null, 2),
          }],
        };
      }

      const fallback = workspaceFallback(workspaces, page.workspace_id);
      const content = await getPageContent(page, fallback);

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            success: true,
            id: page.id,
            title: page.title ?? '(untitled)',
            workspaceId: page.workspace_id,
            content: format === 'html' ? content.html : content.markdown,
          }, null, 2),
        }],
      };
    },
  );
}
