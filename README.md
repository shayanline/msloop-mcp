# msloop-mcp

[![npm version](https://img.shields.io/npm/v/msloop-mcp.svg)](https://www.npmjs.com/package/msloop-mcp)
[![npm downloads](https://img.shields.io/npm/dm/msloop-mcp.svg)](https://www.npmjs.com/package/msloop-mcp)
[![node](https://img.shields.io/node/v/msloop-mcp.svg)](https://www.npmjs.com/package/msloop-mcp)
[![license](https://img.shields.io/npm/l/msloop-mcp.svg)](./LICENSE)

MCP server for Microsoft Loop. No app registration required.

Give any MCP client (Claude, Cursor, Devin, ...) read access to your Microsoft Loop workspaces and pages. It works by reusing your existing Loop web session, the same way [msteams-mcp](https://github.com/shayanline/msteams-mcp) and [msoutlook-mcp](https://github.com/shayanline/msoutlook-mcp) reuse the Teams and Outlook web sessions: you sign in once in a browser, then tokens are cached and refreshed automatically.

## Why

Microsoft does not publish a public API for Loop. This server reuses the Loop web app's own first party client ID, so your access is exactly what your account already has, with no Azure app registration, no admin consent, and no client secrets. Nothing leaves your own machine.

## What it can and cannot do

Loop is built differently from Teams or Outlook. Pages are not REST resources, they are [Fluid Framework](https://fluidframework.com) documents stored in SharePoint Embedded containers, and there is no public content-write API. So this server is a read and discovery tool, not a read-write one:

- **Supported:** list workspaces, list pages in a workspace, read a page's content as Markdown (or raw HTML), search across your Loop files, and an experimental workspace creation.
- **Not supported:** creating or editing page content, real time collaboration. These need the in browser Fluid runtime and have no HTTP API.

Page content is read by asking SharePoint to render the Fluid document to HTML on demand (the `?format=html` export), which is then converted to Markdown. Rich, interactive components (tables, voting, mentions) may render approximately.

## How it works

The Loop web app (`loop.cloud.microsoft`) uses MSAL to store OAuth tokens in the browser. This server:

1. Opens a browser to `loop.cloud.microsoft` via Playwright.
2. Extracts the MSAL tokens from local and session storage, using Loop's own first party client ID (`a187e399-0c36-4b98-8f04-1edc167a0996`). It keeps three: a Substrate token (workspace and page metadata), a SharePoint token (page content), and a Graph token (search).
3. Caches the access tokens, refresh token, and session state in `~/.msloop-mcp-server/` (AES-256-GCM encrypted).
4. Refreshes tokens automatically using the refresh token (HTTP, no browser) or a headless browser as fallback.

## Quick start

```json
{
  "mcpServers": {
    "loop": {
      "command": "npx",
      "args": ["-y", "msloop-mcp@latest"]
    }
  }
}
```

Then run `loop_login` from your MCP client. On first use a browser opens so you can sign in; after that, logins are silent and no browser appears. Do not close the window manually, it closes itself once you are signed in.

## Tools

### Auth

| Tool | Description |
|------|-------------|
| `loop_login` | Sign in to Loop (silent if possible, browser only when needed) |
| `loop_status` | Check authentication status and token validity |
| `loop_logout` | Clear the saved session and tokens |

### Workspaces and pages

| Tool | Description |
|------|-------------|
| `loop_list_workspaces` | List all Loop workspaces you can access, including your personal "My workspace" |
| `loop_list_pages` | List the pages in a workspace (pass a workspace id) |
| `loop_get_page` | Read a page's content as Markdown (or `html`) by page id |
| `loop_search` | Search across your Loop pages and components by keyword (via Microsoft Graph) |
| `loop_create_workspace` | **Experimental.** Create a new shared workspace. Content cannot be created via API |

## Session storage

Session files are stored encrypted in `~/.msloop-mcp-server/`:

- `session-state.json`: Playwright browser session (cookies + localStorage)
- `token-cache.json`: Extracted and cached tokens
- `browser-profile/`: Persistent browser profile for headless refresh

If your session expires, run `loop_login` again.

## Token refresh

Tokens are refreshed automatically:

1. **HTTP refresh** (fast, no browser): uses the cached refresh token with Loop's client ID, one call per resource (Substrate, SharePoint, Graph).
2. **Headless browser refresh**: fallback if HTTP refresh fails; opens a headless browser with the saved profile to silently reacquire tokens.

## Requirements

- Node.js 20+
- A Chromium based browser: Edge or Chrome (detected automatically from system default)
- A Microsoft 365 work or school account with access to Microsoft Loop

## Environment variables

| Variable | Description |
|----------|-------------|
| `MSLOOP_DEBUG=true` | Enable debug logging to stderr |
| `MSLOOP_BROWSER=chrome` | Force a specific browser: `chrome` or `msedge`. If unset, uses the macOS system default; falls back to Chrome on macOS/Linux and Edge on Windows |
| `MSLOOP_CHROME_PROFILE` | Pin a specific Chrome profile dir for cookie import (e.g. `Profile 1`). Defaults to `Default` |
| `MSLOOP_EDGE_PROFILE` | Pin a specific Edge profile dir for cookie import (e.g. `Profile 1`). Defaults to `Default` |
| `MSLOOP_SKIP_COOKIE_IMPORT=true` | Skip importing SSO cookies from your real browser (avoids the one time Keychain/keyring prompt). You sign in once manually; the persistent profile then remembers the session |

## Security notes

- Uses the same auth as the Loop web client, so your access is limited to what your account can do.
- Tokens are encrypted at rest (AES-256-GCM with a machine derived key).
- Uses undocumented internal APIs, which Microsoft may change without notice.

## Acknowledgements

The Loop endpoint mapping (Substrate discovery, the SharePoint multipart "GET via POST" convention, and the HTML export route) was informed by [exec-astraea/loop-migration](https://github.com/exec-astraea/loop-migration) and [Nico De Cleyre's "Microsoft Loop under the hood"](https://www.nicodecleyre.com/blog/2023-04-03-microsoft-loop-under-the-hood/).

## License

MIT. See [LICENSE](./LICENSE).
