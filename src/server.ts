/**
 * MCP server setup — registers all tools and starts the server.
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerAuthTools } from './tools/auth-tools.js';
import { registerWorkspaceTools } from './tools/workspace-tools.js';
import { registerPageTools } from './tools/page-tools.js';
import { registerSearchTools } from './tools/search-tools.js';

export function createServer(): McpServer {
  const server = new McpServer({
    name: 'msloop-mcp',
    version: '0.1.0',
  });

  registerAuthTools(server);
  registerWorkspaceTools(server);
  registerPageTools(server);
  registerSearchTools(server);

  return server;
}
