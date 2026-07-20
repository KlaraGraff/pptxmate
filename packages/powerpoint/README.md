# PPTXMate PowerPoint Add-in

PPTXMate is an open-source PowerPoint AI add-in with an integrated chat panel. It connects to major LLM providers using your own credentials (BYOK) and can read/write presentations through built-in tools, a sandboxed shell, and a virtual filesystem.

## Task-aware context

Requests are routed locally to text, layout, creation, verification, or general workflows. Text-only work uses a compact prompt, an allowlisted tool set, and minimal presentation metadata, while geometry, theme, master, formatting, and raw OOXML are loaded only when the selected workflow needs them. A lightweight slide directory and short previews provide a directory -> selected slides -> requested fields workflow. Stable `slide_id` values remain authoritative when slides are deleted or reordered; only the compact ID/order directory is refreshed, and the target's current position is resolved immediately before a write. Tool results are paginated and bounded, long text can be written back with guarded paragraph/character ranges, only the latest presentation snapshot is retained, and response/context limits use mutation-aware recovery instead of invalidating the conversation or blindly replaying writes. Images can also be pasted directly into the chat composer; supported clipboard images are stored in the session VFS with unique names and can then be read by the agent.

## Stable slide targeting

Call `list_slides` to obtain each stable `slide_id` and the current `directoryVersion`. For tool arguments, pass that value as `directory_version`. A `slide_id` is authoritative across deletion and reordering; `slide_index` is only a zero-based position hint and is re-resolved from the ID immediately before a write. If an index-only request supplies an outdated `directory_version`, the operation is rejected before mutation so the caller can refresh the directory instead of editing the wrong slide.

OOXML-based writes replace the exported slide when importing it back into PowerPoint, which gives the slide a new ID. Their mutation receipt includes `replacementSlideId`; use that value as the authoritative `slide_id` for later operations and refresh `directory_version` from the receipt or `list_slides`.

## Clipboard images

Paste PNG, JPEG, GIF, WebP, or BMP images directly into the chat box. Pasted images appear in the attachment list with unique VFS names; ordinary text paste remains unchanged, and an image-only message can be sent without typing extra text.

## Limits and recovery

When a model reaches its single-response limit, the runtime automatically continues up to two times. If more work remains, send `continue` in the same conversation; the session is not invalidated. Context overflow is handled by compacting old PowerPoint payloads and retrying up to two times. Completed writes are never blindly replayed, and uncertain writes require targeted verification before another mutation.

## Optional CC Switch integration

Local mode exposes `https://localhost:3001/v1` and proxies it to the CC Switch origin at `http://127.0.0.1:15721` by default. CC Switch is a separate application that must be installed and running independently, and account selection remains inside CC Switch. PPTXMate stores no CC Switch credentials.

Configure a custom CC Switch origin or disable the route when starting the local server:

```bash
PPTXMATE_CC_SWITCH_URL=http://127.0.0.1:25721 pnpm dev-server:ppt
PPTXMATE_CC_SWITCH_ENABLED=0 pnpm dev-server:ppt
```

The URL must be an HTTP(S) origin without credentials or a path. The proxy rejects non-local Host headers and cross-origin browser requests. This integration is local-only; the hosted GitHub Pages add-in cannot proxy to a user's localhost.

## Optional macOS lifecycle watcher

Install a user LaunchAgent that starts the local PPTXMate server and its CC Switch route when PowerPoint opens and stops only the server process group it owns when PowerPoint closes. It does not start or stop CC Switch itself, and it leaves an unrelated process already using port `3001` untouched:

```bash
./scripts/install-macos-powerpoint-watcher.sh
```

Watcher-managed sessions can use a custom CC Switch origin or disable the route entirely:

```bash
./scripts/install-macos-powerpoint-watcher.sh --cc-switch-url http://127.0.0.1:25721
./scripts/install-macos-powerpoint-watcher.sh --no-cc-switch
```

Use `--dry-run` to print and validate the generated LaunchAgent without installing it or changing local files.

Uninstall it with:

```bash
./scripts/uninstall-macos-powerpoint-watcher.sh
```

Pass `--remove-logs` to remove only the watcher's managed log files and its empty managed log directory.

The installer discovers repository, Node.js, and pnpm paths dynamically. It does not place personal paths in Git and migrates the earlier local OpenPPT watcher if present.

## Install

Download the hosted [`manifest.prod.xml`](https://klaragraff.github.io/pptxmate/manifest.prod.xml), then follow the instructions for your platform:

### Windows
1. **Insert** → **Add-ins** → **My Add-ins**
2. **Upload My Add-in**
3. Select `manifest.prod.xml`
4. Open the add-in from the ribbon

### macOS
1. Copy `manifest.prod.xml` to:
   `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
2. Restart PowerPoint
3. **Insert** → **Add-ins** → **My Add-ins**
4. Select the add-in

### PowerPoint Web
1. Open [powerpoint.office.com](https://powerpoint.office.com)
2. **Insert** → **Add-ins** → **More Add-ins**
3. **Upload My Add-in**
4. Upload `manifest.prod.xml`

## Tools

| Tool | What it does |
|------|---------------|
| `list_slides` | List up to 25 stable slide IDs/current positions/selection states plus an ID-order version, without reading content |
| `read_slides` | Read up to 25 short plain-text previews by slide ID |
| `screenshot_slide` | Take a screenshot of a slide for visual verification |
| `list_slide_shapes` | List shapes by ID/name/type; geometry is opt-in and paginated |
| `read_slide_texts` | Read normal text boxes as compact plain text; reports omitted table/chart/group containers |
| `read_slide_text` | Read one shape as paginated plain text by default, with range/hash cursors; detailed OOXML is explicit |
| `verify_slides` | Verify layout with a compact issues-only result by default |
| `execute_office_js` | Run raw Office.js inside PowerPoint.run (sandboxed) |
| `edit_slide_text` | Replace/append bounded text, apply a guarded range write, or edit detailed OOXML |
| `update_slide_text` | Batch bounded plain-text changes on one slide with optional stale-text guards |
| `edit_slide_xml` | Edit raw slide XML for advanced layout changes |
| `edit_slide_chart` | Edit chart data and styling in slides |
| `edit_slide_master` | Edit slide master/layout themes |
| `duplicate_slide` | Duplicate an existing slide |
| `read` | Read text files and images from the virtual filesystem |
| `bash` | Run commands in the sandboxed shell |

## Bash custom commands

| Command | What it does |
|---------|---------------|
| `pdf-to-text` | Extract text from PDF files |
| `pdf-to-images` | Render PDF pages to PNG images |
| `docx-to-text` | Extract text from DOCX files |
| `xlsx-to-csv` | Convert uploaded spreadsheet files to CSV |
| `insert-image` | Insert an image into a slide |
| `search-icons` | Search Iconify for vector icons |
| `insert-icon` | Insert an Iconify vector icon into a slide |
| `web-search` | Search the web using configured provider |
| `web-fetch` | Fetch web pages/files into VFS |

Prefer stable IDs for image and icon insertion. These are independent examples using values from the latest `list_slides` result:

```bash
insert-image /home/user/uploads/photo.png --slide-id=256 --directory-version=directory-v1:fnv1a32:7db8a2e1 --x=36 --y=72 --width=240 --unit=pt
insert-icon mdi:check-circle --slide-id=256 --directory-version=directory-v1:fnv1a32:7db8a2e1 --x=300 --y=72 --width=32 --height=32 --color=#16A34A
```

Both commands still accept a legacy 1-based slide number, but stable `--slide-id` targeting is preferred. After either command succeeds, parse its JSON receipt and use `replacementSlideId` plus the returned `directoryVersion` for the next operation.

## Development

```bash
pnpm dev-server:ppt    # Start dev server (https://localhost:3001)
pnpm start:ppt         # Launch PowerPoint with add-in sideloaded
```

## Compatibility identifiers

The internal package name remains `@office-agents/powerpoint`. Browser storage also retains the legacy `openppt-*` keys and IndexedDB name `OpenPPTDB_v1` so existing provider settings, sessions, and uploaded files survive the PPTXMate rename. These identifiers should change only with a tested data migration.

PPTXMate is derived from [Office Agents](https://github.com/hewliyang/office-agents). Repository-wide attribution and license details are in [`NOTICE.md`](../../NOTICE.md) and [`LICENSE`](../../LICENSE).
