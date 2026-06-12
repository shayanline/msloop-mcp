/**
 * Authentication MCP tools.
 */

import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import {
  browserLogin,
  getAuthStatus,
  clearSession,
  getSubstrateToken,
  hasSessionState,
  isSessionLikelyExpired,
} from '../auth/index.js';

/** Minimum minutes remaining before we consider the token "still valid" (skip re-login). */
const TOKEN_VALID_THRESHOLD_MINUTES = 10;

export function registerAuthTools(server: McpServer): void {
  // ── loop_login ─────────────────────────────────────────────────────────
  server.tool(
    'loop_login',
    'Sign in to Microsoft Loop. Tries silently first (no browser); opens a browser only when the session has expired. Set force_new: true to force a full re-login.',
    {
      force_new: z.boolean().optional().describe(
        'Force a full re-login even if a session exists — clears the saved session first. Default: false.',
      ),
    },
    async ({ force_new }) => {
      const forceNew = force_new ?? false;

      if (!forceNew) {
        const existingToken = await getSubstrateToken();
        if (existingToken) {
          const status = getAuthStatus();
          const mins = status.substrateTokenMinutesRemaining ?? 0;
          if (mins >= TOKEN_VALID_THRESHOLD_MINUTES) {
            return {
              content: [{
                type: 'text',
                text: JSON.stringify({
                  success: true,
                  message: `Already authenticated. Token valid for ${mins} more minutes.`,
                  upn: status.upn,
                }, null, 2),
              }],
            };
          }
        }
      }

      if (!forceNew) {
        process.stderr.write(
          '[msloop-mcp] Opening browser. Do NOT close the window — it closes automatically once signed in.\n',
        );
      }

      const result = await browserLogin(forceNew);

      if (!result) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              success: false,
              message: 'Login failed. If you closed the browser window manually, run loop_login again and wait for it to close on its own.',
            }, null, 2),
          }],
        };
      }

      const messages: Record<typeof result.method, string> = {
        'token-cache':    'Already authenticated. Token valid.',
        'headless-sso':   'Login completed silently via SSO. Session has been saved.',
        'headed-browser': 'Login completed successfully. Session has been saved.',
      };

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({ success: true, message: messages[result.method], upn: result.upn }, null, 2),
        }],
      };
    },
  );

  // ── loop_status ────────────────────────────────────────────────────────
  server.tool(
    'loop_status',
    'Check the current Loop authentication status and token validity.',
    {},
    async () => {
      const token = await getSubstrateToken();
      if (!token) {
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({ authenticated: false, message: 'Not authenticated. Run loop_login to sign in.' }, null, 2),
          }],
        };
      }

      const status = getAuthStatus();
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({
            authenticated: true,
            upn: status.upn,
            tenantId: status.tenantId,
            sharePointResource: status.sharePointResource,
            tokens: {
              substrateMinutesRemaining: status.substrateTokenMinutesRemaining,
              sharePointMinutesRemaining: status.sharePointTokenMinutesRemaining,
              graphMinutesRemaining: status.graphTokenMinutesRemaining,
            },
            session: {
              exists: hasSessionState(),
              likelyExpired: isSessionLikelyExpired(),
            },
          }, null, 2),
        }],
      };
    },
  );

  // ── loop_logout ────────────────────────────────────────────────────────
  server.tool(
    'loop_logout',
    'Clear the saved Loop session and tokens. You will need to run loop_login again.',
    {},
    async () => {
      clearSession();
      return {
        content: [{
          type: 'text',
          text: JSON.stringify({ success: true, message: 'Session cleared. Run loop_login to sign in again.' }, null, 2),
        }],
      };
    },
  );
}
