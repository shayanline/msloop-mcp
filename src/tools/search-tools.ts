/**
 * Search MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { searchLoopFiles } from '../api/search.js';

export function registerSearchTools(server: McpServer): void {
  // ── loop_search ──────────────────────────────────────────────────────────
  server.tool(
    'loop_search',
    'Search across your Loop pages and components by keyword (via Microsoft Graph). Returns matching files with their titles, web links, and last-modified dates.',
    {
      query: z.string().min(1).describe('Keywords to search for.'),
      limit: z.number().int().min(1).max(50).optional().describe('Maximum results to return. Default: 25.'),
    },
    async ({ query, limit }) => {
      const hits = await searchLoopFiles(query, limit ?? 25);
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({ count: hits.length, results: hits }, null, 2),
        }],
      };
    },
  );
}
