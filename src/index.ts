#!/usr/bin/env node
/**
 * msloop-mcp — MCP server for Microsoft Loop
 *
 * No app registration required. Uses your existing Loop web session, the same
 * way msteams-mcp and msoutlook-mcp reuse the Teams and Outlook web sessions.
 *
 * Usage: npx msloop-mcp
 */

import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createServer } from './server.js';
import { logger } from './utils/logger.js';

async function main(): Promise<void> {
  const server = createServer();
  const transport = new StdioServerTransport();

  await server.server.connect(transport);
  logger.info('msloop-mcp running on stdio');
}

main().catch(err => {
  logger.error('Fatal error', err);
  process.exit(1);
});
