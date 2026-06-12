# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-06-12

Initial release.

### Added

- Browser-session auth that reuses the Loop web app's own first party client ID, no Azure app registration required. Tokens (Substrate, SharePoint, Graph) are extracted from MSAL local and session storage, cached AES-256-GCM encrypted in `~/.msloop-mcp-server/`, and refreshed over HTTP with a headless browser fallback.
- `loop_login`, `loop_status`, `loop_logout` auth tools.
- `loop_list_workspaces` — list accessible workspaces including the personal "My workspace".
- `loop_list_pages` — list pages in a workspace.
- `loop_get_page` — read a page's content as Markdown or HTML (via SharePoint's on-demand HTML export of the Fluid document).
- `loop_search` — search Loop files by keyword via Microsoft Graph.
- `loop_create_workspace` — experimental shared workspace creation.
- Cross platform SSO cookie import (macOS Keychain, Linux libsecret, Windows DPAPI) for instant silent first login.

[Unreleased]: https://github.com/shayanline/msloop-mcp/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/shayanline/msloop-mcp/releases/tag/v0.1.0
