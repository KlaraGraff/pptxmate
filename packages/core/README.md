# @office-agents/core

`@office-agents/core` is the shared Svelte 5 chat UI layer retained from Office Agents and used by PPTXMate. Its package name remains unchanged for workspace and upstream compatibility.

It re-exports the headless SDK plus the generic chat interface used by the Excel, PowerPoint, and Word add-ins:

- `ChatInterface`
- `FilesPanel`
- `ErrorBoundary`
- app adapter types

## Key pieces

- `src/chat/chat-interface.svelte` — main taskpane chat shell
- `src/chat/chat-controller.ts` — runtime/controller wrapper over `AgentRuntime`
- `src/chat/app-adapter.ts` — app integration contract for Office-specific tools and UI
- `src/chat/settings-panel.svelte` — provider, OAuth, web tools, and skill management
- `src/chat/message-list.svelte` — assistant/user message rendering

## Clipboard images

The chat composer accepts pasted PNG, JPEG, GIF, WebP, and BMP images. Each image is uploaded to the session VFS under a collision-resistant name such as `pasted-image-<id>.png` and appears in the attachment list. A normal text paste is left to the textarea unchanged; the paste event is intercepted only when a supported image is present. A message containing only pasted images can be sent without typing text, in which case the composer supplies a short attachment prompt automatically.

## AppAdapter

Each Office app passes an `AppAdapter` into `ChatInterface` to provide:

- app-specific tools
- system prompt construction
- document identity and metadata
- per-message tool/prompt routing via `toolsForMessage` and `buildSystemPromptForMessage`
- provider-facing context compaction via `transformContext` without rewriting saved chat history
- mutation-aware interruption recovery via `getToolRecoveryInfo` and `normalizeToolArgsForReplay`
- optional Office-specific UI extensions like `ToolExtras`, `HeaderExtras`, and `SelectionIndicator`
- optional link interception via `handleLinkClick`

## Validation

Use the repo-level checks:

```bash
pnpm typecheck
pnpm lint
pnpm build
```
