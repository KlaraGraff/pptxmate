# AGENTS.md

## Project Overview

**PPTXMate** is an open-source PowerPoint AI add-in derived from the Office Agents monorepo. It focuses on token-efficient PowerPoint reading and editing, stable slide targeting, interruption recovery, clipboard image attachments, and optional local integrations. The upstream Excel, Word, SDK, Core, and Bridge packages remain in the workspace so the PowerPoint package can stay compatible with upstream architecture.

The internal `@office-agents/*` package names are intentionally retained. Do not rename them as part of product branding without a coordinated workspace and import migration.

- **@office-agents/sdk** (`packages/sdk/`) — Headless SDK: agent runtime, tools (bash, read), storage, VFS, skills, OAuth, web search/fetch, provider config
- **@office-agents/core** (`packages/core/`) — Svelte chat UI layer: re-exports SDK + ChatInterface, settings panel, sessions, message rendering
- **@office-agents/bridge** (`packages/bridge/`) — Local HTTPS/WebSocket RPC bridge + CLI for talking to a live Office add-in runtime during development
- **@office-agents/excel** (`packages/excel/`) — Excel Add-in: spreadsheet tools, Office.js wrappers, system prompt, cell-range follow mode
- **@office-agents/powerpoint** (`packages/powerpoint/`) — PPTXMate add-in: routed slide reads/writes, stable IDs, OOXML tools, CC Switch proxy, and manifests
- **@office-agents/word** (`packages/word/`) — Word Add-in: document text/structure/OOXML tools, screenshots, Office.js escape hatch

### Key Paths

- `packages/sdk/src/runtime.ts` — `AgentRuntime` class (agent lifecycle, streaming, model resolution)
- `packages/sdk/src/tools/` — Shared tools (`bash.ts`, `read-file.ts`, `types.ts` with `defineTool`)
- `packages/sdk/src/vfs/` — Virtual filesystem + custom commands (`setCustomCommands`)
- `packages/sdk/src/storage/` — IndexedDB sessions, VFS file persistence, skills
- `packages/core/src/chat/` — Svelte chat components and controller (`chat-interface.svelte`, `chat-controller.ts`, `app-adapter.ts`, `settings-panel.svelte`)
- `packages/bridge/src/server.ts` — Local HTTPS/WebSocket bridge server and session registry
- `packages/bridge/src/client.ts` — Add-in bridge client that connects from the Office taskpane to the local bridge
- `packages/bridge/src/cli.ts` — `office-bridge` CLI (`list`, `inspect`, `metadata`, `tool`, `exec`, `events`)
- `packages/excel/src/lib/adapter.ts` — Excel `AppAdapter` (tools, prompt, metadata, follow mode)
- `packages/excel/src/lib/tools/` — Excel-specific tools (`set-cell-range`, `get-cell-ranges`, `eval-officejs`, etc.)
- `packages/powerpoint/src/lib/adapter.ts` — PowerPoint `AppAdapter` (tools, prompt, metadata)
- `packages/powerpoint/src/lib/request-router.ts` — Local task classification and context compaction
- `packages/powerpoint/src/lib/tool-router.ts` — Per-request PowerPoint tool allowlists
- `packages/powerpoint/src/lib/recovery-router.ts` — Mutation-aware recovery scopes and replay normalization
- `packages/powerpoint/src/lib/tools/` — PPT tools (`edit-slide-xml`, `screenshot-slide`, `edit-slide-chart`, etc.)
- `packages/powerpoint/src/lib/pptx/` — OOXML/PPTX helpers, stable slide directory, and text XML utilities
- `packages/powerpoint/cc-switch-proxy.ts` — Optional guarded local CC Switch route
- `scripts/*powerpoint-watcher*` — Optional macOS PowerPoint lifecycle watcher and installer
- `packages/word/src/lib/adapter.ts` — Word `AppAdapter` (tools, prompt, metadata)
- `packages/word/src/lib/tools/` — Word tools (`get-document-text`, `get-document-structure`, `get-paragraph-ooxml`, `screenshot-document`, `execute-office-js`)

## Tech Stack

- **Framework**: Svelte 5
- **Language**: TypeScript
- **Styling**: Tailwind CSS v4 + CSS variables for theming
- **Icons**: Lucide icons (`lucide-svelte`)
- **Build Tool**: Vite 6
- **Office Integration**: Office.js API (`@types/office-js`)
- **LLM Integration**: `@earendil-works/pi-ai` + `@earendil-works/pi-agent-core` (unified LLM & agent API)
- **Virtual Filesystem / Bash**: `just-bash` (in-memory VFS + shell)
- **Dev Server**: Vite dev server with HTTPS
- **Monorepo**: pnpm workspaces

## Key Architecture

### AppAdapter Pattern

Each Office app implements the `AppAdapter` interface from `@office-agents/core`:

```typescript
interface AppAdapter {
  tools: AgentTool[];                               // App-specific tools
  buildSystemPrompt: (skills) => string;            // System prompt
  getDocumentId: () => Promise<string>;             // Unique doc ID for sessions
  getDocumentMetadata?: () => Promise<...>;         // Injected into each prompt
  onToolResult?: (id, result, isError) => void;     // Follow-mode, navigation
  metadataTag?: string;                             // XML tag for metadata (default: "doc_context")
  handleLinkClick?: (context) => "handled" | "default" | Promise<...>;
  ToolExtras?: Component<ToolExtrasProps>;          // Extra UI in tool call blocks
  HeaderExtras?: Component;                         // Extra header controls
  SelectionIndicator?: Component;                   // App-specific selection status UI
  appName?: string;
  appVersion?: string;
  emptyStateMessage?: string;
}
```

The core `ChatInterface` component accepts an adapter and handles all generic chat UI, agent lifecycle, sessions, settings, file uploads, and skills.

### VFS Custom Commands

App-specific VFS commands are registered via `setCustomCommands()` from SDK. PowerPoint adds `insert-image`, `search-icons`, and `insert-icon` to the shared document conversion and web commands. PowerPoint mutations should prefer stable `slide_id` values and use the returned replacement ID after an OOXML re-import.

## Development Commands

```bash
pnpm install             # Install all dependencies
pnpm bridge:serve        # Start the local Office RPC bridge server (https://localhost:4017)
pnpm bridge:stop         # Stop the local Office RPC bridge server
pnpm exec office-bridge list  # List live Office bridge sessions
pnpm dev-server:excel    # Start Excel dev server (https://localhost:3000)
pnpm dev-server:ppt      # Start PowerPoint dev server (https://localhost:3001)
pnpm dev-server:word     # Start Word dev server (https://localhost:3002)
pnpm start:excel         # Launch Excel with add-in sideloaded
pnpm start:ppt           # Launch PowerPoint with add-in sideloaded
pnpm start:word          # Launch Word with add-in sideloaded
pnpm build               # Build all packages
pnpm lint                # Run Biome linter
pnpm format              # Format code with Biome
pnpm typecheck           # TypeScript type checking (all packages)
pnpm check               # Typecheck + lint
pnpm validate            # Validate Office manifests
pnpm licenses:generate   # Refresh THIRD_PARTY_NOTICES.md from the lockfile
./scripts/install-macos-powerpoint-watcher.sh    # Install optional macOS watcher
./scripts/uninstall-macos-powerpoint-watcher.sh  # Remove optional macOS watcher
```

The CC Switch proxy is local-only. CC Switch must be installed, running, and configured separately; the watcher starts and stops PPTXMate's local server, not the CC Switch application. Configure the route with `PPTXMATE_CC_SWITCH_URL` or disable it with `PPTXMATE_CC_SWITCH_ENABLED=0`.

### Office Bridge

During development, the Office taskpane auto-connects to the local bridge client on localhost. Use the bridge to inspect the real Office runtime and run tools against the live add-in:

```bash
pnpm bridge:serve
pnpm bridge:stop
pnpm exec office-bridge list
pnpm exec office-bridge inspect word
pnpm exec office-bridge metadata word
pnpm exec office-bridge tool word get_document_text
pnpm exec office-bridge exec word --code "return { href: window.location.href, title: document.title }"  # unsafe direct eval by default
pnpm exec office-bridge exec word --sandbox --code "const body = context.document.body; body.load('text'); await context.sync(); return body.text;"
pnpm exec office-bridge screenshot word --pages 1 --out page1.png
pnpm exec office-bridge vfs ls word /home/user
pnpm exec office-bridge vfs pull word /home/user/uploads/report.docx ./report.docx
pnpm exec office-bridge vfs push word ./local.txt /home/user/uploads/local.txt
```

`office-bridge exec` runs code with full taskpane/runtime access by default during development. Use `--sandbox` to route through the existing app escape-hatch tool instead.

Use `office-bridge screenshot ... --out file.png` for a simple screenshot-to-local-file workflow, or `office-bridge tool ... --out file.png` for image-returning tool calls. The CLI strips image base64 from printed JSON output to avoid blowing up model context windows.

`pnpm bridge:serve` reuses an already-running healthy bridge server on port `4017` instead of failing with `EADDRINUSE`.

Bridge defaults:

- HTTPS API: `https://localhost:4017`
- WebSocket: `wss://localhost:4017/ws`
- Package docs: `packages/bridge/README.md`

## Code Style

- Formatter/linter: Biome
- No JSDoc comments on functions
- Run `pnpm format` before committing

## Release Workflow

PPTXMate publishes only the PowerPoint add-in from this repository. A push to `main` runs `.github/workflows/deploy-pptxmate-pages.yml`, which tests the PowerPoint package, refreshes third-party notices, builds `packages/powerpoint/dist`, and deploys it to GitHub Pages.

- Production URL: `https://klaragraff.github.io/pptxmate/`
- Public manifest: `https://klaragraff.github.io/pptxmate/manifest.prod.xml`
- Changelog: `packages/powerpoint/CHANGELOG.md`
- Source attribution: `NOTICE.md`

The inherited `scripts/release.mjs` and upstream application tags are not the PPTXMate release path. Do not publish the retained `@office-agents/*` packages to npm or push inherited tags as new PPTXMate releases.

## Configuration Storage

PPTXMate deliberately retains the legacy OpenPPT storage namespace so existing local settings and conversations survive the rename:

| Key | Contents |
| --- | --- |
| `openppt-provider-config` | Provider, model, endpoint, auth method, and local proxy settings |
| `openppt-oauth-credentials` | Provider OAuth access and refresh credentials |
| `openppt-web-config` | Web search/fetch provider configuration |
| `openppt-presentation-id` | Persistent PowerPoint presentation identifier |
| `office-agents-theme` | Shared `"light"` or `"dark"` theme preference |

Session data and VFS files remain in IndexedDB `OpenPPTDB_v1`. The `openppt` and `OpenPPTDB_v1` strings are compatibility identifiers, not stale visible branding. Changing them requires an explicit data migration and tests.

## PowerPoint API Usage

```typescript
await PowerPoint.run(async (context) => {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  return slides.items.map((slide) => slide.id);
});
```

Load only the fields required by the current route. Treat `slide_id` as authoritative and an index as a current directory hint; resolve the ID again immediately before a mutation.

## References

- `packages/bridge/README.md` — bridge usage and CLI docs

- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [PowerPoint JavaScript API](https://learn.microsoft.com/en-us/javascript/api/powerpoint)
- [pi-ai / pi-agent-core](https://github.com/earendil-works/pi-mono)
- [just-bash](https://github.com/nickvdyck/just-bash)
