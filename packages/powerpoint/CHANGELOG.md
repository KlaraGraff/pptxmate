# Changelog

## [Unreleased]

### Added

- **Task-aware context routing** — Text-only requests now use a compact system prompt and minimal presentation metadata instead of loading theme, master, geometry, font, and color data on every turn.
- **Compact text workflow** — Added paginated `read_slide_texts`, plain-text modes for `read_slide_text` / `edit_slide_text`, and batched `update_slide_text` for translation and wording changes without an OOXML round trip.
- **Layered deck reads** — Added a content-free `list_slides` directory and bounded `read_slides` text previews (25 slides, 240 characters per slide by default) before detailed single-slide reads.
- **Guarded paged text writes** — Long text boxes can now be read and written by paragraph/character range using shared cursors and optimistic text hashes, avoiding one oversized model tool call and stale-offset edits.
- **Clipboard image attachments** — PNG, JPEG, GIF, WebP, and BMP images can be pasted directly into the chat composer. Each pasted image receives a unique VFS filename so repeated pastes do not overwrite an earlier attachment; ordinary text paste behavior is unchanged.
- **Stable slide targeting** — Slide reads and writes now prefer PowerPoint's stable `slide_id`, treat `slide_index` as a current-position hint only, and return a lightweight directory version plus relocation/replacement metadata after structural changes.
- **Optional CC Switch integration** — The local PowerPoint server proxies `/v1` to CC Switch, defaults to `http://127.0.0.1:15721`, and supports configurable or disabled routing without storing account credentials in PPTXMate.

### Changed

- **Bounded tool output** — PowerPoint text/tool results now enforce response budgets, expose offsets/continuation hints, and keep verification output issues-only unless full shape geometry is explicitly requested.
- **Tool-schema routing** — Text and discovery requests expose only their allowlisted tools; full Office.js/OOXML schemas are loaded only for explicit design, structure, creation, verification, or specialized container work.
- **Context recovery** — Only the latest `<ppt_context>` snapshot is retained, old PowerPoint tool payloads are compacted before provider calls, and length/context-limit responses are automatically continued up to two times. Mutation receipts now preserve completed writes, block exact replays, and require a matching targeted verifier before continuing uncertain text/layout/structure writes; arbitrary mutations remain fail-closed.
- **Structure refresh** — Adding, deleting, moving, duplicating, or re-importing a slide refreshes only the ID/order directory. Index-only calls with a stale directory stop before mutation, while ID-bound calls relocate to the target's current position without rereading fonts, colors, geometry, or slide content.
- **Office requirement** — The manifest now requires PowerPointApi 1.10 so null-object text-frame reads are reliable across every shape type.

## [0.0.6] - 2026-05-12

### Changed

- **Upgrade `pi-ai` / `pi-agent-core`** — Migrated from `@mariozechner/pi-ai` (deprecated) to `@earendil-works/pi-ai` `^0.74.0`. Adds `gpt-5.5` and other latest models to the picker.

## [0.0.5] - 2026-04-17

### Changed

- **Co-located command prompt snippets** — PowerPoint-specific VFS commands (`insert-image`, `search-icons`, `insert-icon`) now use `DescribedCommand` with co-located `promptSnippet`. System prompt assembles command docs dynamically from snippets instead of hard-coding them.
- **Upgrade `pi-ai` / `pi-agent-core`** — Bumped to `^0.67.6` for the latest provider, streaming, and agent runtime improvements.

## [0.0.4] - 2026-03-19

### Changed

- **Svelte migration** — Ported `@office-agents/core` chat UI layer from React to Svelte 5.

## [0.0.3] - 2026-03-15

### Features

- **Dev bridge integration** — In development mode the taskpane auto-connects to the local Office bridge, enabling CLI-driven tool calls, screenshots, VFS access, and live inspection.
- **Files panel** — New "Files" tab in the chat header lets you browse, preview, download, and delete VFS files.

### Fixes

- **`btoa`/`atob` in `execute_office_js`** — Base64 helpers are now available inside the Office.js sandbox.
- **CSS source path** — Fixed `streamdown` Tailwind `@source` path after monorepo restructure.

## [0.0.2] - 2026-03-08

### Fixes

- **PDF commands** — Fixed `pdf-to-text` and `pdf-to-images` consuming the PDF file data on first use, causing subsequent calls to fail with "The object can not be cloned".

## [0.0.1] - 2026-03-08

Initial release with AI chat interface, multi-provider LLM support (BYOK), PowerPoint slide read/write tools, OOXML/PPTX editing, and CORS proxy configuration.
