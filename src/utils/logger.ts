/**
 * Minimal logger that writes to stderr so MCP stdout stays clean.
 */

const isDebug = process.env.MSLOOP_DEBUG === 'true';

export const logger = {
  info: (msg: string, ...args: unknown[]) => {
    process.stderr.write(`[msloop-mcp] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
  debug: (msg: string, ...args: unknown[]) => {
    if (!isDebug) return;
    process.stderr.write(`[msloop-mcp:debug] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
  warn: (msg: string, ...args: unknown[]) => {
    process.stderr.write(`[msloop-mcp:warn] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
  error: (msg: string, ...args: unknown[]) => {
    process.stderr.write(`[msloop-mcp:error] ${msg}${args.length ? ' ' + JSON.stringify(args) : ''}\n`);
  },
};
