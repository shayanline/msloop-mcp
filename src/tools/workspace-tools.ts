/**
 * Workspace MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { listWorkspaces, createWorkspace } from '../api/loop.js';

export function registerWorkspaceTools(server: McpServer): void {
  // ── loop_list_workspaces ─────────────────────────────────────────────────
  server.tool(
    'loop_list_workspaces',
    'List all Microsoft Loop workspaces you can access, including your personal "My workspace". Returns each workspace id and title.',
    {},
    async () => {
      const workspaces = await listWorkspaces();
      const simplified = workspaces.map(w => ({
        id: w.id,
        title: w.title ?? '(untitled)',
        isPersonal: w.isPersonal ?? false,
        hasStorage: !!w.mfs_info?.pod_id,
      }));
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({ count: simplified.length, workspaces: simplified }, null, 2),
        }],
      };
    },
  );

  // ── loop_create_workspace (experimental) ─────────────────────────────────
  server.tool(
    'loop_create_workspace',
    'EXPERIMENTAL: Create a new shared Loop workspace. Workspace creation is only partially documented and the new workspace may need a moment to finish provisioning in the Loop app. Creating page content is not supported.',
    {
      name: z.string().min(1).describe('Display name for the new workspace.'),
    },
    async ({ name }) => {
      try {
        const created = await createWorkspace(name);
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: true,
              message: `Requested creation of workspace "${name}". It may take a moment to appear in Loop.`,
              workspace: created,
            }, null, 2),
          }],
        };
      } catch (err) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: false,
              message: `Workspace creation failed: ${err instanceof Error ? err.message : String(err)}`,
            }, null, 2),
          }],
        };
      }
    },
  );
}
