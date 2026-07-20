# PPTXMate

PPTXMate is an open-source PowerPoint AI assistant and Office add-in for reading, editing, translating, and automating PPT/PPTX presentations with less model context and fewer unnecessary tokens.

PPTXMate 是一个开源 PowerPoint AI 插件，支持按需读取幻灯片、文本修改与翻译、稳定页面定位、图片粘贴和长任务恢复。

> PPTXMate is an independent project. It is not affiliated with or endorsed by Microsoft.

## Highlights

- **Task-aware context routing**: text, layout, creation, and verification requests expose only the prompt, tools, and presentation fields they need.
- **Layered, bounded reads**: start with a compact slide directory, preview selected slides, and load detailed text, geometry, formatting, or OOXML only on demand.
- **Stable slide identity**: writes target PowerPoint `slide_id` values rather than stale page numbers after deletion or reordering.
- **Limit recovery**: bounded outputs, transcript compaction, automatic continuation, mutation receipts, and inspect-before-retry recovery keep the same conversation usable.
- **Clipboard images**: paste PNG, JPEG, GIF, WebP, or BMP images directly into the chat composer.
- **BYOK providers**: use API keys, supported OAuth flows, or compatible custom endpoints.
- **Optional CC Switch integration**: route the local `/v1` endpoint through the account currently selected in CC Switch.
- **Optional macOS lifecycle watcher**: start the local PPTXMate route when PowerPoint opens and stop it when PowerPoint closes.

PowerPoint implementation details and the complete tool list are documented in [`packages/powerpoint`](./packages/powerpoint).

## Install The Hosted Add-in

The GitHub Pages workflow publishes the task pane and public manifest at:

[`https://klaragraff.github.io/pptxmate/manifest.prod.xml`](https://klaragraff.github.io/pptxmate/manifest.prod.xml)

Download that manifest and sideload it in PowerPoint. Platform-specific steps are in the [PowerPoint package README](./packages/powerpoint/README.md#install).

The hosted version supports normal BYOK and compatible HTTPS endpoints. CC Switch uses a localhost route and therefore requires the local mode below.

## Local Development

Requirements:

- Node.js 20 or later
- pnpm 11.9.0 or later
- Microsoft PowerPoint desktop for add-in testing

```bash
git clone https://github.com/KlaraGraff/pptxmate.git
cd pptxmate
pnpm install
pnpm start:ppt
```

The PowerPoint development server is available at `https://localhost:3001`.

## Optional CC Switch Route

PPTXMate can proxy its local OpenAI-compatible endpoint to CC Switch. CC Switch is a separate application that must be installed, running, and configured independently; PPTXMate never stores CC Switch account credentials.

The default route is:

```text
https://localhost:3001/v1 -> http://127.0.0.1:15721/v1
```

In PPTXMate Settings, configure a custom endpoint with the API type and model supported by CC Switch, then use `https://localhost:3001/v1` as the Base URL. Change or disable the route before starting the local server with:

```bash
PPTXMATE_CC_SWITCH_URL=http://127.0.0.1:25721 pnpm dev-server:ppt
PPTXMATE_CC_SWITCH_ENABLED=0 pnpm dev-server:ppt
```

Switch accounts in CC Switch itself. Subsequent PPTXMate requests use the account selected there. The route rejects cross-origin browser requests and non-local host headers.

## Optional macOS Auto Start/Stop

The macOS watcher keeps no model or account credentials. It watches for the `Microsoft PowerPoint` process, starts the local PPTXMate development server and CC Switch route when PowerPoint opens, and terminates only the server process group it owns when PowerPoint closes. It does not start, stop, or switch accounts in the separate CC Switch application. If port `3001` already belongs to another process, the watcher leaves that process untouched.

```bash
./scripts/install-macos-powerpoint-watcher.sh
```

To persist a custom CC Switch address, or disable that route for watcher-managed sessions, provide the setting while installing:

```bash
./scripts/install-macos-powerpoint-watcher.sh --cc-switch-url http://127.0.0.1:25721
./scripts/install-macos-powerpoint-watcher.sh --no-cc-switch
```

Preview the generated LaunchAgent without installing or changing local files with `./scripts/install-macos-powerpoint-watcher.sh --dry-run`.

Remove it with:

```bash
./scripts/uninstall-macos-powerpoint-watcher.sh
```

Use `./scripts/uninstall-macos-powerpoint-watcher.sh --remove-logs` to remove the watcher's own log files as well.

The installer discovers the current repository, Node.js, and pnpm paths and generates a user-specific LaunchAgent outside the repository. It also migrates the earlier local OpenPPT watcher when present.

## Repository Structure

PPTXMate keeps the shared SDK and UI packages from the upstream Office Agents monorepo for compatibility:

| Package | Purpose |
| --- | --- |
| [`@office-agents/powerpoint`](./packages/powerpoint) | PPTXMate PowerPoint add-in, routing, slide tools, and manifests |
| [`@office-agents/core`](./packages/core) | Shared Svelte chat UI and clipboard attachment handling |
| [`@office-agents/sdk`](./packages/sdk) | Agent runtime, context recovery, providers, storage, VFS, and sandbox |
| [`@office-agents/bridge`](./packages/bridge) | Optional local runtime inspection bridge |
| [`@office-agents/excel`](./packages/excel) | Retained upstream Excel add-in package |
| [`@office-agents/word`](./packages/word) | Retained upstream Word add-in package |

Internal package names remain unchanged to minimize divergence from upstream. The PowerPoint storage identifiers `openppt-*` and `OpenPPTDB_v1` are also retained so existing local settings and conversations survive the product rename. References to the earlier OpenPPT LaunchAgent label in the watcher scripts exist only to migrate or remove that legacy agent.

## Validation

```bash
pnpm test
pnpm typecheck
pnpm lint
pnpm validate
pnpm build
```

## Upstream And License

PPTXMate is derived from [hewliyang/office-agents](https://github.com/hewliyang/office-agents) and retains its Git history. See [NOTICE.md](./NOTICE.md) for attribution.

Released under the [MIT License](./LICENSE).
