# Agent Guidelines for Loop MCP

This document captures project knowledge to help AI agents work effectively with this codebase.

## Repository

- **Repository**: https://github.com/shayanline/msloop-mcp
- **Package**: [`msloop-mcp`](https://www.npmjs.com/package/msloop-mcp)
- **Install**: `npx -y msloop-mcp@latest`, or clone the repo, `npm install && npm run build`, then point your MCP client to `dist/index.js`.
- **Sibling projects**: [`msteams-mcp`](https://github.com/shayanline/msteams-mcp) and [`msoutlook-mcp`](https://github.com/shayanline/msoutlook-mcp). This project follows the same auth and session pattern; `msoutlook-mcp` is the closest template.

## Project Overview

An MCP server that gives AI assistants read and discovery access to Microsoft Loop. Microsoft publishes no Loop API, so this reuses the Loop web app's own first party client ID and the internal APIs the web client calls, with tokens extracted from a browser session. The browser is only used for initial login and silent refresh fallback, everything else is direct HTTP.

## The central constraint

Loop is not a REST resource model like mail or calendar. Workspaces are SharePoint Embedded containers, and pages are Fluid Framework documents (ops plus snapshots). There is no public content-write API and content cannot be materialised over plain HTTP without the Fluid client runtime.

So this server is **read and discovery only**:

- Metadata (workspaces, pages) comes from the Substrate Loop API.
- Page content is read by asking SharePoint to render the Fluid document to HTML on demand (`?format=html`), then converting to Markdown. Lossy for rich components.
- Writing or editing page content is **out of scope** and not achievable here.

## Architecture

```
src/
  index.ts              Entry point, runs the MCP server on stdio
  server.ts             createServer() — registers all tool groups
  constants.ts          Client ID, endpoints, scopes, timeouts
  auth/
    index.ts            getSubstrateToken / getSharePointToken / getGraphToken, status
    session-store.ts    AES-256-GCM encrypted token + session storage in ~/.msloop-mcp-server/
    token-extractor.ts  Pure: select Substrate/SharePoint/Graph tokens from MSAL entries by audience
    token-refresh.ts    HTTP refresh, one OAuth2 call per resource
    browser-login.ts    Playwright login at loop.cloud.microsoft (headless first)
  browser/
    cookie-import.ts    Cross-platform SSO cookie import (Keychain / libsecret / DPAPI)
  api/
    client.ts           substrateGet/Post, graphGet/Post, sharePointGetText
    loop.ts             discover() (merge /workspaces + /recent + /deltasync), listPages, createWorkspace
    pages.ts            getPageContent — HTML export then htmlToMarkdown
    search.ts           Graph /search/query for .loop / .fluid files
  utils/
    http.ts             Bearer headers, retry, the SharePoint multipart "GET via POST" builder
    parsers.ts          Pure: decodePodId, itemIdFromPageId, hostFromSiteUrl, slugify, htmlToMarkdown
    logger.ts           stderr logger (MSLOOP_DEBUG=true)
  types/loop.ts         Substrate API shapes
```

## Auth and endpoints

- **Client ID**: `a187e399-0c36-4b98-8f04-1edc167a0996` (Loop web app, also the SPE container type ID). Public SPA, so token refresh requires the `Origin: https://loop.cloud.microsoft` header (Azure AD returns AADSTS9002327 without it).
- **Three tokens**, selected from the MSAL cache by audience, since Loop may store tokens in localStorage or sessionStorage (we read both):
  - Substrate (`substrate.office.com`) — workspace and page metadata.
  - SharePoint (`{tenant}.sharepoint.com`) — Fluid snapshots and page content.
  - Graph (`graph.microsoft.com`) — file search.
- **Discovery**: `GET https://substrate.office.com/recommended/api/v1.1/loop/{workspaces,recent,deltasync}`. No single endpoint is complete, so all three are merged and deduped by id, following `next_page_link`.
- **Page content**: `GET https://{spHost}/_api/v2.0/drives/{driveId}/items/{itemId}/content?format=html&ump=1`, fetched via the multipart "GET via POST" convention (`X-HTTP-Method-Override: GET` inside a `multipart/form-data` body). Coordinates come from the page's `sharepoint_info.site_url` and `onedrive_info.drive_id`, with the item id taken from the page id after the last `_`; the workspace `mfs_info.pod_id` (base64 `…|host|driveId|itemId`) is the fallback.
- **Workspace creation (experimental)**: `POST https://substrate.office.com/speedway/v1.0/workspaceGroups`.

## Implementation patterns

1. **Credential encryption** at rest (AES-256-GCM, machine-derived scrypt key, 0o600 files), auto-migrating any legacy plaintext.
2. **HTTP-first token refresh** per resource with a headless browser fallback; first login always needs a browser.
3. **Pure logic is isolated and unit-tested** (`parsers.ts`, `token-extractor.ts`, the multipart builder). Browser and network code is excluded from coverage.
4. **Tools return JSON text content** and never throw raw, mirroring the sibling projects.

## Development

- `npm run typecheck`, `npm run build`, `npm test` (Vitest). All must pass before a release.
- Real end to end auth cannot be exercised in CI, it requires an interactive `loop_login` against a real tenant. The pure logic is covered by tests instead.
- Release process: see `RELEASING.md` (GitHub Release triggers npm Trusted Publishing).

## Tools

`loop_login`, `loop_status`, `loop_logout`, `loop_list_workspaces`, `loop_list_pages`, `loop_get_page`, `loop_search`, `loop_create_workspace` (experimental).
