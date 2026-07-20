# @office-agents/sdk

Headless SDK for building AI-powered Microsoft Office Add-ins. Provides an agent runtime, tool system, virtual filesystem, session storage, and multi-provider LLM integration — all running in the browser.

PPTXMate retains the upstream `@office-agents/sdk` package name for compatibility. This repository does not republish the inherited package under that npm scope.

> **Browser-only** — this package targets browser environments (Office Add-ins, SPAs). It uses IndexedDB, localStorage, and the DOM.

## Install

```bash
npm install @office-agents/sdk
```

## Overview

The SDK is organized into several modules:

| Module | Description |
| --- | --- |
| **AgentContext** | Owns the VFS, bash shell, static files, skill files, and custom commands for one agent instance |
| **Runtime** | `AgentRuntime` — manages agent lifecycle, streaming, model resolution, sessions |
| **Tools** | `defineTool`, `createBashTool`, `createReadTool` — define and register tools for the agent |
| **VFS** | In-memory virtual filesystem with bash shell (`just-bash`) — file uploads, custom commands |
| **Storage** | IndexedDB-backed session persistence, VFS file storage, skill file storage |
| **Provider Config** | Multi-provider LLM configuration (OpenAI, Anthropic, Google, etc.) |
| **OAuth** | PKCE OAuth flow helpers for provider authentication |
| **Skills** | Installable skill system (prompt snippets + files mounted at `/home/skills/`) |
| **Web** | Web search and fetch with pluggable providers |
| **Sandbox** | `SES`-based sandboxed JavaScript evaluation |

## Quick Start

### 1. Define a tool

```typescript
import { defineTool, toolSuccess } from "@office-agents/sdk";
import { Type } from "@sinclair/typebox";

const greetTool = defineTool({
  name: "greet",
  label: "Greet",
  description: "Greet someone by name",
  parameters: Type.Object({
    name: Type.String({ description: "Name to greet" }),
  }),
  execute: async (toolCallId, params) => {
    return toolSuccess({ message: `Hello, ${params.name}!` });
  },
});
```

### 2. Create a RuntimeAdapter

The `RuntimeAdapter` interface connects your app to the agent runtime:

```typescript
import type { RuntimeAdapter } from "@office-agents/sdk";

const adapter: RuntimeAdapter = {
  tools: [greetTool],

  buildSystemPrompt: (skills, commandSnippets) => {
    let prompt = "You are a helpful assistant. Use tools when appropriate.";
    if (commandSnippets.length > 0) {
      prompt += "\n\nAvailable commands:\n" + commandSnippets.join("\n");
    }
    return prompt;
  },

  getDocumentId: async () => {
    return "my-document-id"; // unique ID for session scoping
  },

  // Optional: provide static files to mount in the VFS (e.g., API docs)
  staticFiles: {
    "/home/user/docs/api-reference.d.ts": "declare const MyApi: any;",
  },

  // Optional: register custom bash commands
  customCommands: (ns) => ({
    commands: [/* CustomCommand instances */],
    promptSnippets: ["Use `my-cmd <arg>` to do something"],
  }),

  // Optional: inject context into each message
  getDocumentMetadata: async () => ({
    metadata: { title: "My Document", sheets: ["Sheet1"] },
  }),

  // Optional: react to tool results (e.g., navigate to a cell)
  onToolResult: (toolCallId, result, isError) => {
    console.log("Tool result:", result);
  },

  // Optional: scope storage to your app
  storageNamespace: {
    dbName: "MyAppDB",
    dbVersion: 1,
    localStoragePrefix: "my-app",
    documentSettingsPrefix: "my-app",
  },
};
```

### 3. Create a context and initialize the runtime

`AgentContext` owns all per-instance state: the virtual filesystem, bash shell, static files, skill files, and custom commands. The runtime applies adapter config to the context during `init()`.

```typescript
import { AgentContext, AgentRuntime } from "@office-agents/sdk";

// Option A: let runtime.init() apply adapter.staticFiles / adapter.customCommands
const ctx = new AgentContext({
  namespace: adapter.storageNamespace,
});

// Option B: pass everything upfront (app entrypoints typically do this)
const ctx = new AgentContext({
  namespace: adapter.storageNamespace,
  staticFiles: adapter.staticFiles,
  customCommands: adapter.customCommands,
});

const runtime = new AgentRuntime(adapter, ctx);

// Subscribe to state changes
runtime.subscribe((state) => {
  console.log("Messages:", state.messages.length);
  console.log("Streaming:", state.isStreaming);
});

// Initialize (applies adapter config, loads saved provider config, restores session + VFS)
await runtime.init();

// Send a message
await runtime.sendMessage("Hello, who are you?");
```

### 4. Virtual filesystem & bash

The VFS and bash shell are accessed through `AgentContext`:

```typescript
import { AgentContext } from "@office-agents/sdk";

const ctx = new AgentContext({
  staticFiles: {
    "/home/user/docs/guide.txt": "Getting started...",
  },
});

// Write files to VFS
await ctx.writeFile("/home/user/uploads/data.csv", "name,age\nAlice,30");
// Relative paths resolve to /home/user/uploads/
await ctx.writeFile("notes.txt", "some notes");

// Read files
const content = await ctx.readFile("/home/user/uploads/data.csv");

// List uploads
const uploads = await ctx.listUploads(); // ["data.csv", "notes.txt"]

// Execute bash commands
const bash = ctx.bash;
const result = await bash.exec("ls /home/user/uploads/");

// Snapshot and restore VFS state (used for session persistence)
const snapshot = await ctx.snapshotVfs();  // excludes /home/skills/
await ctx.restoreVfs(snapshot);

// Update skill files in-place (preserves user uploads)
await ctx.setSkillFiles({
  "/home/skills/analysis/SKILL.md": new TextEncoder().encode("# Analysis skill"),
});

// Update static files in-place (preserves user uploads)
await ctx.setStaticFiles({
  "/home/user/docs/new-guide.txt": "Updated guide",
});
```

### 5. Built-in tools

The SDK provides factory functions that create tools bound to an `AgentContext`:

```typescript
import { createBashTool, createReadTool } from "@office-agents/sdk";

const ctx = new AgentContext();

// Tools that operate on the context's VFS
const bashTool = createBashTool(ctx);
const readTool = createReadTool(ctx);

// Or use tools as a function of context in your adapter:
const adapter: RuntimeAdapter = {
  tools: (ctx) => [createBashTool(ctx), createReadTool(ctx), greetTool],
  // ...
};
```

### 6. Provider configuration

All config functions take a `StorageNamespace` to scope localStorage keys:

```typescript
import { loadSavedConfig, saveConfig, type StorageNamespace } from "@office-agents/sdk";

const ns: StorageNamespace = {
  dbName: "MyAppDB",
  dbVersion: 1,
  localStoragePrefix: "my-app",
  documentSettingsPrefix: "my-app",
};

// Load saved config from localStorage
const config = loadSavedConfig(ns);

// Save a new config
saveConfig(ns, {
  provider: "openai",
  apiKey: "sk-...",
  model: "gpt-4o",
  useProxy: false,
  proxyUrl: "",
  thinking: "none",
  followMode: true,
  expandToolCalls: false,
});
```

### 7. Sessions

Session and VFS persistence functions are scoped by namespace and document ID:

```typescript
import {
  createSession,
  listSessions,
  getSession,
  saveSession,
  deleteSession,
  loadVfsFiles,
  saveVfsFiles,
  type StorageNamespace,
} from "@office-agents/sdk";

const ns: StorageNamespace = {
  dbName: "MyAppDB",
  dbVersion: 1,
  localStoragePrefix: "my-app",
  documentSettingsPrefix: "my-app",
};

const docId = "my-document-id";

// Create and manage sessions
const session = await createSession(ns, docId, "My Chat");
const sessions = await listSessions(ns, docId);
await deleteSession(ns, session.id);

// Persist VFS files per session
const snapshot = await ctx.snapshotVfs();
await saveVfsFiles(ns, session.id, snapshot);
const restored = await loadVfsFiles(ns, session.id);
await ctx.restoreVfs(restored);
```

### 8. Skills

Skills are installable packages containing a `SKILL.md` with frontmatter and optional supporting files. They're mounted at `/home/skills/<name>/` in the VFS:

```typescript
import {
  addSkill,
  getInstalledSkills,
  removeSkill,
  syncSkillsToVfs,
  buildSkillsPromptSection,
  AgentContext,
  type StorageNamespace,
} from "@office-agents/sdk";

const ns: StorageNamespace = { /* ... */ };
const ctx = new AgentContext({ namespace: ns });

// Install a skill from files
await addSkill(ns, ctx, [
  {
    path: "SKILL.md",
    data: "---\nname: data-analysis\ndescription: Data analysis best practices\n---\n\n# Data Analysis\n\nWhen analyzing data, always start by summarizing the dataset...",
  },
]);

// List installed skills
const skills = await getInstalledSkills(ns);

// Build the skills section for the system prompt
const promptSection = buildSkillsPromptSection(skills);

// Sync all skill files into the VFS (preserves uploads)
await syncSkillsToVfs(ns, ctx);

// Uninstall
await removeSkill(ns, ctx, "data-analysis");
```

### 9. Custom bash commands

Register custom commands that appear in the bash shell and generate prompt snippets:

```typescript
import { getSharedCustomCommands, type CustomCommandsResult, type StorageNamespace } from "@office-agents/sdk";

function getCustomCommands(ns: StorageNamespace): CustomCommandsResult {
  // getSharedCustomCommands provides built-in commands like
  // pdf-to-text, docx-to-text, xlsx-to-csv, web-search, web-fetch
  return getSharedCustomCommands({ ns });
}

const ctx = new AgentContext({
  customCommands: getCustomCommands,
});

// Command snippets are derived from registered commands
const snippets = ctx.commandSnippets; // string[]
```

## API Reference

### AgentContext

- **`new AgentContext(opts?)`** — Create a context with optional `namespace`, `staticFiles`, `skillFiles`, `customCommands`.
- **`.vfs`** — The `InMemoryFs` instance (lazy-initialized).
- **`.bash`** — The `Bash` shell instance (lazy-initialized).
- **`.commandSnippets`** — Prompt snippets derived from registered custom commands.
- **`.writeFile(path, content)`** / **`.readFile(path)`** / **`.readFileBuffer(path)`** / **`.fileExists(path)`** / **`.deleteFile(path)`** — File operations (relative paths resolve to `/home/user/uploads/`).
- **`.listUploads()`** — List files in `/home/user/uploads/` (excludes `.keep`).
- **`.snapshotVfs()`** — Snapshot all files except `/home/skills/`.
- **`.restoreVfs(files)`** — Reset and restore from snapshot + static/skill overlays.
- **`.setStaticFiles(files)`** / **`.setSkillFiles(files)`** — Update mounted files in-place (preserves user data).
- **`.setCustomCommands(factory)`** — Replace custom command factory (rebuilds bash).
- **`.reset()`** — Discard VFS and bash (next access re-initializes from static + skill files).

### Runtime

- **`new AgentRuntime(adapter, context)`** — Create a runtime with an adapter and context.
- **`.init()`** — Apply adapter config, load saved provider config, restore session and VFS.
- **`.sendMessage(text, images?)`** — Send a user message and stream the response.
- **`.applyConfig(config)`** — Apply a provider configuration.
- **`.newSession()`** / **`.switchSession(id)`** / **`.deleteCurrentSession()`** — Session management.
- **`.uploadFiles(files)`** / **`.removeUpload(name)`** — Manage VFS uploads.
- **`.installSkill(files)`** / **`.uninstallSkill(name)`** — Install/remove skills.
- **`.subscribe(listener)`** — Subscribe to `RuntimeState` changes.
- **`.dispose()`** — Abort streaming and clean up.

### RuntimeAdapter

- **`tools`** — `AgentTool[]` or `(ctx: AgentContext) => AgentTool[]`.
- **`buildSystemPrompt(skills, commandSnippets)`** — Build the system prompt.
- **`toolsForMessage?(message, ctx, info)`** — Return a per-message tool allowlist; `info.priorUserMessages` comes from the active session.
- **`buildSystemPromptForMessage?(message, skills, commands, info)`** — Build a task-aware prompt using the same session-scoped preparation info.
- **`getDocumentId()`** — Return a unique document ID for session scoping.
- **`staticFiles?`** — Files to mount in VFS (applied during `init()`).
- **`customCommands?`** — Custom command factory (applied during `init()`).
- **`getDocumentMetadata?`** — Inject context metadata into each prompt; receives the user request, abort signal, and message preparation info.
- **`transformContext?`** — Compact the provider-facing transcript without mutating persisted session history.
- **`getToolRecoveryInfo?`** / **`normalizeToolArgsForReplay?`** — Classify mutations and normalize replay identity for interruption-safe recovery.
- **`onToolResult?`** — React to tool results.
- **`metadataTag?`** — XML tag name for metadata injection.
- **`storageNamespace?`** — Override default storage namespace.

### Per-message routing and context

`toolsForMessage` and `buildSystemPromptForMessage` run for each real user request, so an app can classify that request and expose only the relevant tool schemas and instructions. `MessagePreparationInfo.priorUserMessages` contains prior real requests from the active session and excludes hidden runtime continuation/recovery messages.

`transformContext(messages, signal, info)` runs immediately before a provider request. It may return a compact provider-facing transcript based on the model context window, system-prompt size, and recovery attempt. The returned transcript does not replace the messages persisted in the session, so UI history and subsequent user turns remain intact.

### Interruption and limit recovery

`getToolRecoveryInfo(toolName, args)` returns a `ToolRecoveryInfo` with:

- `effect`: `read`, `write`, or `unknown`.
- `mutationKind`: `text`, `layout`, `structure`, or `arbitrary` for a possible write.
- `verificationKinds`: mutation kinds that a read can verify.
- `scope`: compact non-content identity such as `slide_id` and `shape_id`.

Write tools should include a structured `mutationState` in their result: `not_started` means no mutation occurred and retry is allowed, `completed` means the receipt is authoritative and replay is blocked, and `uncertain` means the write may have landed and targeted verification is required before another write. `normalizeToolArgsForReplay` can make equivalent calls share the same replay identity.

For a response stopped by the model's output limit, the runtime automatically continues up to two times. For a context overflow, it calls `transformContext` and automatically retries or enters mutation-aware recovery up to two times. If either limit remains after those attempts, the session is preserved: the user can send “continue” or a narrower request in the same conversation instead of creating a new one.

### Tools

- **`defineTool(config)`** — Create a typed tool with name, description, parameters (TypeBox schema), and execute function.
- **`toolSuccess(data)`** / **`toolError(message)`** / **`toolText(text)`** — Build tool results.
- **`createBashTool(ctx)`** — Create a bash tool bound to a context.
- **`createReadTool(ctx)`** — Create a file-read tool bound to a context.

### Storage

- **`createSession(ns, workbookId, name?)`** / **`saveSession(ns, session)`** / **`deleteSession(ns, id)`** — Session CRUD.
- **`listSessions(ns, workbookId)`** / **`getSession(ns, id)`** — Query sessions.
- **`getOrCreateCurrentSession(ns, workbookId)`** — Get or create the current session for a document.
- **`saveVfsFiles(ns, sessionId, files)`** / **`loadVfsFiles(ns, sessionId)`** — Persist VFS files per session.
- **`getOrCreateDocumentId(ns)`** — Get or create a persistent document ID.

### Provider Config

- **`loadSavedConfig(ns)`** / **`saveConfig(ns, config)`** — Read/write provider settings from localStorage.
- **`buildCustomModel(config)`** — Build a model instance from a custom base URL config.
- **`applyProxyToModel(model, config)`** — Apply proxy URL to a model.

### Skills

- **`addSkill(ns, ctx, files)`** — Install a skill from file inputs.
- **`removeSkill(ns, ctx, name)`** — Uninstall a skill.
- **`getInstalledSkills(ns)`** — List installed skills with metadata.
- **`syncSkillsToVfs(ns, ctx)`** — Sync all skill files into the VFS.
- **`buildSkillsPromptSection(skills)`** — Build the prompt section listing available skills.
- **`parseSkillMeta(content)`** — Parse SKILL.md frontmatter.

### OAuth

- **`generatePKCE()`** — Generate PKCE code verifier + challenge.
- **`buildAuthorizationUrl(provider, ...)`** / **`exchangeOAuthCode(...)`** / **`refreshOAuthToken(...)`** — Full OAuth flow.

### Web

- **`searchWeb(query, options)`** / **`searchImages(query, options)`** — Web search with pluggable providers.
- **`fetchWeb(url, options)`** — Fetch and extract web page content.

## Used By

- **[@office-agents/excel](https://github.com/hewliyang/office-agents/tree/main/packages/excel)** — Excel Add-in with AI chat
- **[PPTXMate PowerPoint Add-in](https://github.com/KlaraGraff/pptxmate/tree/main/packages/powerpoint)** — Token-efficient PowerPoint AI add-in using the retained `@office-agents/powerpoint` workspace package name
- **[@office-agents/word](https://github.com/hewliyang/office-agents/tree/main/packages/word)** — Word Add-in with AI chat

## License

MIT
