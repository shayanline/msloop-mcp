/**
 * Shared constants used across the codebase.
 */

// ─────────────────────────────────────────────────────────────────────────────
// OAuth / Auth
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Microsoft's own first-party application ID for the Loop web app
 * (loop.cloud.microsoft). This is the "Loop / Chapter 5 Fluid App" client ID,
 * which is also the SharePoint Embedded container type ID Loop content lives in.
 * No app registration required — same pattern as msteams-mcp / msoutlook-mcp,
 * which reuse Teams' and Outlook's own first-party client IDs.
 */
export const LOOP_CLIENT_ID = 'a187e399-0c36-4b98-8f04-1edc167a0996';

/** OAuth token endpoint for work/school accounts (used as default when no tenant ID). */
export const TOKEN_ENDPOINT_TEMPLATE = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token';

/** Resource scope for the Substrate Loop API (workspace + page metadata). */
export const SUBSTRATE_SCOPE = 'https://substrate.office.com/.default';

/** Microsoft Graph scope (used for file search and metadata). */
export const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';

/** The Loop web app URL we open for login and token acquisition. */
export const LOOP_URL = 'https://loop.cloud.microsoft/';

/** The web origin the Loop SPA presents — required as the Origin header for SPA token refresh. */
export const LOOP_ORIGIN = 'https://loop.cloud.microsoft';

// ─────────────────────────────────────────────────────────────────────────────
// API Base URLs
// ─────────────────────────────────────────────────────────────────────────────

/** Substrate host that serves the internal Loop metadata API. */
export const SUBSTRATE_BASE = 'https://substrate.office.com';

/** Substrate Loop API (workspaces, recent, deltasync) — version 1.1. */
export const LOOP_API_BASE = `${SUBSTRATE_BASE}/recommended/api/v1.1/loop`;

/** Substrate Speedway API (workspace group creation) — version 1.0. */
export const SPEEDWAY_BASE = `${SUBSTRATE_BASE}/speedway/v1.0`;

/** Microsoft Graph base. */
export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ─────────────────────────────────────────────────────────────────────────────
// Token / Refresh
// ─────────────────────────────────────────────────────────────────────────────

/** Refresh tokens ~55 minutes before expiry (access tokens last ~1 hour). */
export const TOKEN_REFRESH_BUFFER_MS = 55 * 60 * 1000;

/** How long to wait for the user to complete browser login (ms). */
export const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;

// ─────────────────────────────────────────────────────────────────────────────
// HTTP
// ─────────────────────────────────────────────────────────────────────────────

/** User-Agent presented to Loop/Substrate/SharePoint APIs so requests look like the web client. */
export const LOOP_USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0';
