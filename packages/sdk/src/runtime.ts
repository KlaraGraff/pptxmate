import {
  Agent,
  type AgentEvent,
  type AgentMessage,
  type ThinkingLevel as AgentThinkingLevel,
  type AgentTool,
  type ToolExecutionMode,
} from "@earendil-works/pi-agent-core";
import {
  type Api,
  type AssistantMessage,
  getModel,
  getModels,
  getProviders,
  isContextOverflow,
  type Message,
  type Model,
  streamSimple,
} from "@earendil-works/pi-ai";
import type { AgentContext, StorageNamespace } from "./context";
import {
  agentMessagesToChatMessages,
  type ChatMessage,
  deriveStats,
  extractPartsFromAssistantMessage,
  generateId,
  type RuntimeUsageCarry,
  type SessionStats,
  stripEnrichment,
} from "./message-utils";
import {
  loadOAuthCredentials,
  refreshOAuthToken,
  saveOAuthCredentials,
} from "./oauth";
import {
  applyProxyToModel,
  buildCustomModel,
  loadSavedConfig,
  type ProviderConfig,
  saveConfig,
  type ThinkingLevel,
} from "./provider-config";
import {
  addSkill,
  getInstalledSkills,
  removeSkill,
  type SkillMeta,
  syncSkillsToVfs,
} from "./skills";
import {
  type ChatSession,
  createSession,
  deleteSession,
  getOrCreateCurrentSession,
  getSession,
  listSessions,
  loadVfsFiles,
  saveSession,
  saveVfsFiles,
} from "./storage";
import type { CustomCommandsResult } from "./vfs/custom-commands";

export interface RuntimeAdapter {
  tools: AgentTool[] | ((ctx: AgentContext) => AgentTool[]);
  /** Optional per-message tool allowlist used to reduce schemas and enforce routing. */
  toolsForMessage?: (
    userMessage: string,
    ctx: AgentContext,
    info: MessagePreparationInfo,
  ) => AgentTool[];
  buildSystemPrompt: (skills: SkillMeta[], commandSnippets: string[]) => string;
  /** Optional task-aware prompt rebuilt immediately before each user message. */
  buildSystemPromptForMessage?: (
    userMessage: string,
    skills: SkillMeta[],
    commandSnippets: string[],
    info?: MessagePreparationInfo,
  ) => string;
  getDocumentId: () => Promise<string>;
  getDocumentMetadata?: (request?: {
    userMessage: string;
    messageCount: number;
    signal?: AbortSignal;
    info?: MessagePreparationInfo;
  }) => Promise<{
    metadata: object;
    nameMap?: Record<number, string>;
  } | null>;
  metadataHistory?: "all" | "latest";
  /** Optional app-specific classification used to make interruption recovery safe. */
  getToolRecoveryInfo?: (
    toolName: string,
    args: unknown,
  ) => ToolRecoveryInfo | undefined;
  /** Normalize semantically equivalent arguments before replay-key hashing. */
  normalizeToolArgsForReplay?: (toolName: string, args: unknown) => unknown;
  /** Optional app-specific wording for a hidden recovery instruction. */
  buildRecoveryPrompt?: (input: RecoveryPromptInput) => string;
  /** Serialize document tools so a write cannot race another write during recovery. */
  toolExecution?: ToolExecutionMode;
  /** Optional app-specific context pruning before each provider request. */
  transformContext?: (
    messages: AgentMessage[],
    signal?: AbortSignal,
    info?: {
      contextWindow: number;
      systemPromptChars: number;
      recoveryAttempt: number;
    },
  ) => Promise<AgentMessage[]>;
  onToolResult?: (toolCallId: string, result: string, isError: boolean) => void;
  metadataTag?: string;
  staticFiles?: Record<string, string>;
  customCommands?: (ns: StorageNamespace) => CustomCommandsResult;
  storageNamespace?: Partial<StorageNamespace>;
}

export interface MessagePreparationInfo {
  /** Previous real user requests from the active session, oldest first. */
  priorUserMessages: readonly string[];
}

export type ToolEffect = "read" | "write" | "unknown";
export type RecoveryKind = "text" | "layout" | "structure" | "arbitrary";

export interface ToolRecoveryInfo {
  effect: ToolEffect;
  /** Kind of mutation a write may perform. Defaults to arbitrary. */
  mutationKind?: RecoveryKind;
  /** Mutation kinds a read result can verify after the model observes it. */
  verificationKinds?: RecoveryKind[];
  /** Small, non-content scope such as slide_index and shape_id. */
  scope?: Record<string, string | number | boolean>;
}

export interface RecoveryMutation {
  toolCallId: string;
  toolName: string;
  effect: Exclude<ToolEffect, "read">;
  kind: RecoveryKind;
  status: "completed" | "uncertain";
  scope?: Record<string, string | number | boolean>;
}

export interface RecoveryPromptInput {
  reason: "overflow" | "resume_after_abort";
  attempt: number;
  mutations: readonly RecoveryMutation[];
}

export interface UploadedFile {
  name: string;
  size: number;
}

export interface RuntimeState {
  messages: ChatMessage[];
  isStreaming: boolean;
  error: string | null;
  providerConfig: ProviderConfig | null;
  sessionStats: SessionStats;
  currentSession: ChatSession | null;
  sessions: ChatSession[];
  nameMap: Record<number, string>;
  uploads: UploadedFile[];
  isUploading: boolean;
  skills: SkillMeta[];
  vfsInvalidatedAt: number;
}

type StateListener = (state: RuntimeState) => void;

const INITIAL_STATS: SessionStats = { ...deriveStats([]), contextWindow: 0 };

interface ToolJournalEntry {
  toolCallId: string;
  toolName: string;
  info: ToolRecoveryInfo;
  status: "started" | "completed" | "failed";
  callKey: string;
  replayKeys: Set<string>;
  args: unknown;
  admitted: boolean;
}

type JsonRecord = Record<string, unknown>;

function isJsonRecord(value: unknown): value is JsonRecord {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

function stableValue(value: unknown): string {
  if (value === null || typeof value !== "object") {
    return JSON.stringify(value) ?? String(value);
  }
  if (Array.isArray(value)) return `[${value.map(stableValue).join(",")}]`;
  const record = value as Record<string, unknown>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableValue(record[key])}`)
    .join(",")}}`;
}

function toolCallKey(toolName: string, args: unknown): string {
  const input = `${toolName}:${stableValue(args)}`;
  let hash = 2166136261;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return `${toolName}:${(hash >>> 0).toString(16)}`;
}

function toolResultText(result: unknown): string {
  if (typeof result === "string") return result;
  if (
    result &&
    typeof result === "object" &&
    "content" in result &&
    Array.isArray((result as { content?: unknown }).content)
  ) {
    return (
      result as { content: Array<{ type?: string; text?: string }> }
    ).content
      .filter((part) => part.type === "text")
      .map((part) => part.text ?? "")
      .join("\n");
  }
  return "";
}

function parseJsonRecord(value: string): JsonRecord | undefined {
  try {
    const parsed: unknown = JSON.parse(value);
    return isJsonRecord(parsed) ? parsed : undefined;
  } catch {
    return undefined;
  }
}

function toolResultRecords(result: unknown): JsonRecord[] {
  const records: JsonRecord[] = [];
  if (isJsonRecord(result) && !("content" in result)) records.push(result);

  const text = toolResultText(result).trim();
  if (!text) return records;
  const complete = parseJsonRecord(text);
  if (complete) {
    records.push(complete);
    return records;
  }

  // Bash custom commands can surround their structured result with normal
  // stdout/stderr. A standalone JSON line still carries the tool contract.
  for (const line of text.split(/\r?\n/)) {
    const record = parseJsonRecord(line.trim());
    if (record) records.push(record);
  }
  return records;
}

function toolResultFailed(result: unknown, isError: boolean): boolean {
  if (isError) return true;
  return toolResultRecords(result).some(
    (record) => record.success === false || typeof record.error === "string",
  );
}

function mutationDidNotStart(records: readonly JsonRecord[]): boolean {
  let completedFallback: boolean | undefined;
  for (let index = records.length - 1; index >= 0; index--) {
    const record = records[index];
    const state = record.mutationState ?? record.mutation_state;
    if (typeof state === "string") {
      const normalized = state.trim().toLowerCase().replace(/[ -]+/g, "_");
      if (normalized === "not_started") return true;
      // A structured state is authoritative over the legacy boolean, even if
      // a transitional producer emits contradictory fields.
      if (normalized === "uncertain" || normalized === "completed") {
        return false;
      }
    }
    const completed = record.mutationCompleted ?? record.mutation_completed;
    if (typeof completed === "boolean" && completedFallback === undefined) {
      completedFallback = completed;
    }
  }
  return completedFallback === false;
}

function mergeRecoveryScopes(
  previous: ToolRecoveryInfo["scope"],
  derived: ToolRecoveryInfo["scope"],
): ToolRecoveryInfo["scope"] {
  if (!derived) return previous;
  const scope = { ...previous, ...derived };
  if (derived.slide_id !== undefined) delete scope.slide_index;
  if (derived.slide_ids !== undefined) delete scope.slide_indices;
  return Object.keys(scope).length > 0 ? scope : undefined;
}

const RESULT_SCOPE_ALIASES: Record<string, readonly string[]> = {
  slide_id: [
    "slide_id",
    "slideId",
    "current_slide_id",
    "currentSlideId",
    "_modifiedSlideId",
    "modifiedSlideId",
    "replacement_slide_id",
    "replacementSlideId",
    "new_slide_id",
    "newSlideId",
  ],
  slide_index: [
    "slide_index",
    "slideIndex",
    "current_slide_index",
    "currentSlideIndex",
    "_modifiedSlide",
    "modifiedSlide",
  ],
  directory_version: ["directory_version", "directoryVersion"],
  original_slide_id: ["original_slide_id", "originalSlideId"],
  replacement_slide_id: ["replacement_slide_id", "replacementSlideId"],
  new_slide_id: ["new_slide_id", "newSlideId"],
  source_slide_id: ["source_slide_id", "sourceSlideId"],
  shape_id: ["shape_id", "shapeId"],
  slide_ids: ["slide_ids", "slideIds"],
  slide_indices: ["slide_indices", "slideIndices"],
  shape_ids: ["shape_ids", "shapeIds", "updates"],
};

function scopeDerivedFromResultRecord(
  scope: ToolRecoveryInfo["scope"],
  record: JsonRecord,
): ToolRecoveryInfo["scope"] {
  if (!scope) return undefined;
  const entries = Object.entries(scope).filter(([key]) => {
    const aliases = RESULT_SCOPE_ALIASES[key] ?? [key];
    return aliases.some((alias) => alias in record);
  });
  return entries.length > 0 ? Object.fromEntries(entries) : undefined;
}

function argsWithResultSlideId(
  args: unknown,
  scope: ToolRecoveryInfo["scope"],
): unknown | undefined {
  if (!isJsonRecord(args) || typeof scope?.slide_id !== "string") {
    return undefined;
  }
  return { ...args, slide_id: scope.slide_id };
}

function compactRecoveryScopeValue(value: string | number | boolean): string {
  return String(value).replace(/\s+/g, " ").slice(0, 80);
}

function sanitizeRecoveryInfo(info: ToolRecoveryInfo): ToolRecoveryInfo {
  const effect: ToolEffect = ["read", "write", "unknown"].includes(info.effect)
    ? info.effect
    : "unknown";
  const scopeEntries = Object.entries(info.scope ?? {})
    .filter((entry): entry is [string, string | number | boolean] =>
      ["string", "number", "boolean"].includes(typeof entry[1]),
    )
    .slice(0, 8)
    .map(([key, value]) => [
      key.replace(/[^a-zA-Z0-9_.-]/g, "").slice(0, 40),
      compactRecoveryScopeValue(value),
    ]);
  return {
    effect,
    mutationKind: ["text", "layout", "structure", "arbitrary"].includes(
      info.mutationKind ?? "",
    )
      ? info.mutationKind
      : effect === "read"
        ? undefined
        : "arbitrary",
    verificationKinds: Array.from(new Set(info.verificationKinds ?? [])).filter(
      (kind): kind is RecoveryKind =>
        ["text", "layout", "structure", "arbitrary"].includes(kind),
    ),
    scope:
      scopeEntries.length > 0 ? Object.fromEntries(scopeEntries) : undefined,
  };
}

function thinkingLevelToAgent(level: ThinkingLevel): AgentThinkingLevel {
  return level === "none" ? "off" : level;
}

export class AgentRuntime {
  readonly context: AgentContext;

  private agent: Agent | null = null;
  private config: ProviderConfig | null = null;
  private pendingConfig: ProviderConfig | null = null;
  private streamingMessageId: string | null = null;
  private isStreaming = false;
  private documentId: string | null = null;
  private currentSessionId: string | null = null;
  private sessionLoaded = false;
  private followMode = true;
  private skills: SkillMeta[] = [];
  private continuationNeeded = false;
  private continuationReason: "length" | "overflow" | null = null;
  private autoContinuationCount = 0;
  private contextRecoveryAttempt = 0;
  private toolJournal = new Map<string, ToolJournalEntry>();
  private pendingRecovery = new Map<string, RecoveryMutation>();
  private completedRecoveryKeys = new Set<string>();
  private verifiedRecoveryIds = new Set<string>();
  private recoveryVerificationRequired = false;
  private manualAbortPending = false;
  private lastRunSucceeded = false;
  private sendGeneration = 0;
  private pendingSendController: AbortController | null = null;
  private uploadOperationCount = 0;
  private uploadQueue: Promise<void> = Promise.resolve();

  private adapter: RuntimeAdapter;
  private listeners: Set<StateListener> = new Set();
  private state: RuntimeState;

  private get ns(): StorageNamespace {
    return this.context.namespace;
  }

  private get tools(): AgentTool[] {
    return typeof this.adapter.tools === "function"
      ? this.adapter.tools(this.context)
      : this.adapter.tools;
  }

  constructor(adapter: RuntimeAdapter, context: AgentContext) {
    this.adapter = adapter;
    this.context = context;

    const saved = loadSavedConfig(this.ns);
    const validConfig =
      saved?.provider && saved?.apiKey && saved?.model ? saved : null;
    this.followMode = validConfig?.followMode ?? true;
    this.state = {
      messages: [],
      isStreaming: false,
      error: null,
      providerConfig: validConfig,
      sessionStats: INITIAL_STATS,
      currentSession: null,
      sessions: [],
      nameMap: {},
      uploads: [],
      isUploading: false,
      skills: [],
      vfsInvalidatedAt: 0,
    };
  }

  getState(): RuntimeState {
    return this.state;
  }

  subscribe(listener: StateListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  private emit() {
    for (const listener of this.listeners) {
      listener(this.state);
    }
  }

  private update(partial: Partial<RuntimeState>) {
    this.state = { ...this.state, ...partial };
    this.emit();
  }

  private bumpVfs() {
    this.update({ vfsInvalidatedAt: Date.now() });
  }

  private updateMessages(
    updater: (messages: ChatMessage[]) => ChatMessage[],
    extra?: Partial<RuntimeState>,
  ) {
    this.state = {
      ...this.state,
      messages: updater(this.state.messages),
      ...extra,
    };
    this.emit();
  }

  setAdapter(adapter: RuntimeAdapter) {
    this.adapter = adapter;
  }

  getAvailableProviders(): string[] {
    return getProviders();
  }

  getModelsForProvider(provider: string): Model<Api>[] {
    try {
      return (getModels as (p: string) => Model<Api>[])(provider);
    } catch {
      return [];
    }
  }

  private async getActiveApiKey(config: ProviderConfig): Promise<string> {
    if (config.authMethod !== "oauth") {
      return config.apiKey;
    }
    const creds = loadOAuthCredentials(this.ns, config.provider);
    if (!creds) return config.apiKey;
    if (Date.now() < creds.expires) {
      return creds.access;
    }
    const refreshed = await refreshOAuthToken(
      config.provider,
      creds.refresh,
      config.proxyUrl,
      config.useProxy,
    );
    saveOAuthCredentials(this.ns, config.provider, refreshed);
    return refreshed.access;
  }

  private classifyTool(
    toolName: string,
    args: unknown,
  ): ToolRecoveryInfo | undefined {
    if (!this.adapter.getToolRecoveryInfo) return undefined;
    try {
      return sanitizeRecoveryInfo(
        this.adapter.getToolRecoveryInfo(toolName, args) ?? {
          effect: "unknown",
        },
      );
    } catch {
      // A classifier failure must fail closed: an unclassified tool may mutate
      // the document, so recovery should inspect it before replaying anything.
      return { effect: "unknown" };
    }
  }

  private recordToolStart(
    toolCallId: string,
    toolName: string,
    args: unknown,
  ): void {
    const info = this.classifyTool(toolName, args);
    if (!info) return;
    if (this.toolJournal.size >= 128) {
      const oldestRead = Array.from(this.toolJournal.entries()).find(
        ([, entry]) => entry.info.effect === "read",
      )?.[0];
      if (oldestRead) this.toolJournal.delete(oldestRead);
    }
    const callKey = this.getToolCallKey(toolName, args);
    this.toolJournal.set(toolCallId, {
      toolCallId,
      toolName,
      info,
      status: "started",
      callKey,
      replayKeys: new Set([callKey]),
      args,
      admitted: false,
    });
  }

  private recordToolEnd(
    toolCallId: string,
    toolName: string,
    args: unknown,
    result: unknown,
    isError: boolean,
  ): void {
    const existing = this.toolJournal.get(toolCallId);
    const callKey = this.getToolCallKey(toolName, args);
    const entry =
      existing ??
      ({
        toolCallId,
        toolName,
        info: this.classifyTool(toolName, args) ?? { effect: "unknown" },
        status: "started",
        callKey,
        replayKeys: new Set([callKey]),
        args,
        admitted: false,
      } satisfies ToolJournalEntry);
    const resultRecords = toolResultRecords(result);
    let resultScope: ToolRecoveryInfo["scope"];
    for (const record of resultRecords) {
      const resultInfo = this.classifyTool(toolName, record);
      resultScope = mergeRecoveryScopes(
        resultScope,
        scopeDerivedFromResultRecord(resultInfo?.scope, record),
      );
    }
    entry.info = {
      ...entry.info,
      scope: mergeRecoveryScopes(entry.info.scope, resultScope),
    };
    const resultTargetArgs = argsWithResultSlideId(entry.args, resultScope);
    if (resultTargetArgs !== undefined) {
      entry.replayKeys.add(this.getToolCallKey(toolName, resultTargetArgs));
    }

    const failed = toolResultFailed(result, isError);
    const didNotStart = failed && mutationDidNotStart(resultRecords);
    const uncertainFailure =
      failed && !didNotStart && entry.info.effect !== "read" && entry.admitted;
    entry.status = failed ? "failed" : "completed";
    this.toolJournal.set(toolCallId, entry);

    if (didNotStart) {
      this.pendingRecovery.delete(toolCallId);
      this.verifiedRecoveryIds.delete(toolCallId);
      this.refreshRecoveryVerificationRequirement();
    } else if (uncertainFailure) {
      this.pendingRecovery.set(toolCallId, {
        toolCallId,
        toolName,
        effect: entry.info.effect === "write" ? "write" : "unknown",
        kind: entry.info.mutationKind ?? "arbitrary",
        status: "uncertain",
        scope: entry.info.scope,
      });
      this.refreshRecoveryVerificationRequirement();
    }
    if (
      !failed &&
      (entry.info.verificationKinds?.length ?? 0) > 0 &&
      this.recoveryVerificationRequired
    ) {
      for (const mutation of this.pendingRecovery.values()) {
        if (!this.verificationCoversMutation(entry.info, mutation)) continue;
        this.verifiedRecoveryIds.add(mutation.toolCallId);
      }
    }
    if (this.pendingRecovery.has(toolCallId)) {
      const pending = this.pendingRecovery.get(toolCallId)!;
      pending.status = failed ? "uncertain" : "completed";
      pending.scope = entry.info.scope;
      this.pendingRecovery.set(toolCallId, pending);
      if (!failed) {
        for (const key of entry.replayKeys) this.completedRecoveryKeys.add(key);
      }
      this.refreshRecoveryVerificationRequirement();
    }
  }

  private currentRecoveryMutations(): RecoveryMutation[] {
    return Array.from(this.toolJournal.values())
      .filter(
        (entry) =>
          entry.info.effect !== "read" &&
          entry.status !== "failed" &&
          (entry.status !== "started" || entry.admitted),
      )
      .map((entry) => ({
        toolCallId: entry.toolCallId,
        toolName: entry.toolName,
        effect: entry.info.effect as Exclude<ToolEffect, "read">,
        kind: entry.info.mutationKind ?? "arbitrary",
        status: entry.status === "completed" ? "completed" : "uncertain",
        scope: entry.info.scope,
      }));
  }

  private captureRecoveryMutations(): void {
    for (const mutation of this.currentRecoveryMutations()) {
      const previous = this.pendingRecovery.get(mutation.toolCallId);
      this.pendingRecovery.set(mutation.toolCallId, {
        ...mutation,
        status:
          previous?.status === "completed" || mutation.status === "completed"
            ? "completed"
            : "uncertain",
      });
      if (mutation.status === "completed") {
        const entry = this.toolJournal.get(mutation.toolCallId);
        if (entry) {
          for (const key of entry.replayKeys) {
            this.completedRecoveryKeys.add(key);
          }
        }
      }
    }
    this.refreshRecoveryVerificationRequirement();
  }

  private refreshRecoveryVerificationRequirement(): void {
    this.recoveryVerificationRequired = Array.from(
      this.pendingRecovery.values(),
    ).some(
      (mutation) =>
        mutation.status === "uncertain" || mutation.effect === "unknown",
    );
  }

  private commitObservedVerifications(): void {
    for (const toolCallId of this.verifiedRecoveryIds) {
      this.pendingRecovery.delete(toolCallId);
      this.toolJournal.delete(toolCallId);
    }
    this.verifiedRecoveryIds.clear();
    this.refreshRecoveryVerificationRequirement();
  }

  private verificationCoversMutation(
    verification: ToolRecoveryInfo,
    mutation: RecoveryMutation,
  ): boolean {
    if (!verification.verificationKinds?.includes(mutation.kind)) return false;
    const readScope = verification.scope;
    const mutationScope = mutation.scope;
    if (!readScope) return true;
    if (!mutationScope) return false;

    const same = (key: string) =>
      readScope[key] !== undefined &&
      mutationScope[key] !== undefined &&
      String(readScope[key]) === String(mutationScope[key]);
    const listContains = (listKey: string, itemKey: string) => {
      const list = readScope[listKey];
      const item = mutationScope[itemKey];
      if (list === undefined || item === undefined) return false;
      return String(list)
        .split(",")
        .map((value) => value.trim())
        .includes(String(item));
    };

    if (mutationScope.shape_id !== undefined) {
      if (readScope.shape_id === undefined || !same("shape_id")) return false;
      if (readScope.slide_index !== undefined) return same("slide_index");
      if (readScope.slide_id !== undefined) return same("slide_id");
      return true;
    }
    return (
      same("slide_index") ||
      same("slide_id") ||
      listContains("slide_indices", "slide_index") ||
      listContains("slide_ids", "slide_id")
    );
  }

  private recoveryMutations(): RecoveryMutation[] {
    return Array.from(this.pendingRecovery.values()).slice(-16);
  }

  private settleResolvedRecovery(): void {
    const unresolved = Array.from(this.pendingRecovery.values()).some(
      (mutation) =>
        mutation.status === "uncertain" || mutation.effect === "unknown",
    );
    if (unresolved) return;
    this.pendingRecovery.clear();
    this.completedRecoveryKeys.clear();
    this.verifiedRecoveryIds.clear();
    this.recoveryVerificationRequired = false;
  }

  private defaultRecoveryPrompt(
    reason: RecoveryPromptInput["reason"],
    attempt: number,
  ): string {
    const receipts = this.recoveryMutations()
      .map((mutation) => {
        const scope = mutation.scope
          ? ` (${Object.entries(mutation.scope)
              .slice(0, 8)
              .map(
                ([key, value]) =>
                  `${key.replace(/[^a-zA-Z0-9_.-]/g, "").slice(0, 40)}=${compactRecoveryScopeValue(value)}`,
              )
              .join(", ")})`
          : "";
        return `- ${mutation.status} ${mutation.kind}: ${mutation.toolName}${scope}`;
      })
      .join("\n");
    const needsVerification = this.recoveryMutations().some(
      (mutation) =>
        mutation.status === "uncertain" || mutation.effect === "unknown",
    );
    const needsManualConfirmation = this.recoveryMutations().some(
      (mutation) =>
        (mutation.status === "uncertain" || mutation.effect === "unknown") &&
        mutation.kind === "arbitrary",
    );
    const nextStep = needsManualConfirmation
      ? "Use the smallest targeted reads to inspect the current document. An arbitrary mutation cannot be cleared by a narrow automatic verifier, so remain read-only and explain what must be confirmed before further document writes."
      : needsVerification
        ? "Before any new write, use the smallest targeted read or verification tool to inspect the current document, then apply only changes that are still missing."
        : "Continue from the retained successful tool results without repeating them.";
    return `<runtime_recovery>\nA previous document run was ${
      reason === "overflow" ? "interrupted by a context limit" : "stopped"
    } after possible document mutations (recovery attempt ${attempt}).\n${receipts}\nTreat completed tool results as authoritative and do not replay those writes. Treat uncertain writes as possibly applied. ${nextStep} Do not include code, XML, or full text in this recovery step.\n</runtime_recovery>`;
  }

  private buildRecoveryPrompt(
    reason: RecoveryPromptInput["reason"],
    attempt: number,
  ): string {
    const input: RecoveryPromptInput = {
      reason,
      attempt,
      mutations: this.recoveryMutations(),
    };
    try {
      const custom = this.adapter.buildRecoveryPrompt?.(input)?.trim();
      if (!custom) return this.defaultRecoveryPrompt(reason, attempt);
      if (/^<runtime_recovery>[\s\S]*<\/runtime_recovery>$/.test(custom)) {
        return custom;
      }
      return `<runtime_recovery>\n${custom}\n</runtime_recovery>`;
    } catch {
      return this.defaultRecoveryPrompt(reason, attempt);
    }
  }

  private removeOverflowEnvelope(promptMessageIndex: number): void {
    const messages = this.agent?.state.messages;
    if (!messages) return;
    let removedFailure = false;
    for (let i = messages.length - 1; i >= promptMessageIndex; i--) {
      const message = messages[i];
      if (message.role !== "assistant") {
        if (removedFailure) break;
        continue;
      }
      const assistant = message as AssistantMessage;
      const emptyFailure =
        (assistant.stopReason === "error" ||
          assistant.stopReason === "aborted") &&
        assistant.content.every(
          (block) => block.type === "text" && block.text.length === 0,
        );
      if (
        !isContextOverflow(assistant, this.state.sessionStats.contextWindow) &&
        !emptyFailure
      ) {
        break;
      }
      this.addUsageToPrompt(promptMessageIndex, assistant);
      messages.splice(i, 1);
      removedFailure = true;
    }
  }

  private addUsageToPrompt(
    promptMessageIndex: number,
    assistant: AssistantMessage,
  ): void {
    const messages = this.agent?.state.messages;
    const prompt = messages?.[promptMessageIndex];
    if (!prompt || prompt.role !== "user" || !assistant.usage) return;
    const carrier = prompt as typeof prompt & {
      _runtimeUsage?: RuntimeUsageCarry;
    };
    const carry: RuntimeUsageCarry = carrier._runtimeUsage ?? {
      inputTokens: 0,
      outputTokens: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalCost: 0,
      lastInputTokens: 0,
    };
    carry.inputTokens += assistant.usage.input;
    carry.outputTokens += assistant.usage.output;
    carry.cacheRead += assistant.usage.cacheRead;
    carry.cacheWrite += assistant.usage.cacheWrite;
    carry.totalCost += assistant.usage.cost.total;
    carry.lastInputTokens =
      assistant.usage.input +
      assistant.usage.cacheRead +
      assistant.usage.cacheWrite;
    carrier._runtimeUsage = carry;
  }

  private prepareRecovery(
    promptMessageIndex: number,
    reason: RecoveryPromptInput["reason"],
    attempt: number,
  ): string {
    this.captureRecoveryMutations();
    this.removeOverflowEnvelope(promptMessageIndex);
    return this.buildRecoveryPrompt(reason, attempt);
  }

  private resetRecoveryState(): void {
    this.toolJournal.clear();
    this.pendingRecovery.clear();
    this.completedRecoveryKeys.clear();
    this.verifiedRecoveryIds.clear();
    this.recoveryVerificationRequired = false;
    this.manualAbortPending = false;
  }

  private isReplayGuarded(toolName: string, args: unknown): boolean {
    return this.completedRecoveryKeys.has(this.getToolCallKey(toolName, args));
  }

  private getToolCallKey(toolName: string, args: unknown): string {
    try {
      const normalized =
        this.adapter.normalizeToolArgsForReplay?.(toolName, args) ?? args;
      return toolCallKey(toolName, normalized);
    } catch {
      return toolCallKey(toolName, args);
    }
  }

  private markToolAdmitted(toolCallId: string): void {
    const entry = this.toolJournal.get(toolCallId);
    if (entry) entry.admitted = true;
  }

  private getMessagePreparationInfo(agent: Agent): MessagePreparationInfo {
    const priorUserMessages = agent.state.messages
      .filter((message) => message.role === "user")
      .map((message) =>
        stripEnrichment(message.content, this.adapter.metadataTag).trim(),
      )
      .filter((message) => message.length > 0);
    return { priorUserMessages };
  }

  private handleAgentEvent = async (event: AgentEvent) => {
    console.log("[Runtime] Agent event:", event.type, event);
    switch (event.type) {
      case "message_start": {
        if (event.message.role === "assistant") {
          const id = generateId();
          this.streamingMessageId = id;
          const parts = extractPartsFromAssistantMessage(event.message);
          const chatMessage: ChatMessage = {
            id,
            role: "assistant",
            parts,
            timestamp: event.message.timestamp,
          };
          this.updateMessages((msgs) => [...msgs, chatMessage]);
        }
        break;
      }
      case "turn_end": {
        this.commitObservedVerifications();
        break;
      }
      case "message_update": {
        if (event.message.role === "assistant" && this.streamingMessageId) {
          const streamId = this.streamingMessageId;
          this.updateMessages((msgs) => {
            const messages = [...msgs];
            const idx = messages.findIndex((m) => m.id === streamId);
            if (idx !== -1) {
              const parts = extractPartsFromAssistantMessage(
                event.message,
                messages[idx].parts,
              );
              messages[idx] = { ...messages[idx], parts };
            }
            return messages;
          });
        }
        break;
      }
      case "message_end": {
        if (event.message.role === "assistant") {
          const assistantMsg = event.message as AssistantMessage;
          const overflow = isContextOverflow(
            assistantMsg,
            this.state.sessionStats.contextWindow,
          );
          this.lastRunSucceeded =
            !overflow && assistantMsg.stopReason === "stop";
          if (overflow) {
            this.continuationNeeded = true;
            this.continuationReason = "overflow";
            this.agent?.abort();
          } else if (assistantMsg.stopReason === "length") {
            const hasToolCall = assistantMsg.content.some(
              (block) => block.type === "toolCall",
            );
            if (!hasToolCall && this.autoContinuationCount < 2) {
              this.autoContinuationCount++;
              this.agent?.followUp({
                role: "user",
                content:
                  "<runtime_continue>Continue exactly where the previous response stopped. Do not repeat completed content.</runtime_continue>",
                timestamp: Date.now(),
              });
            } else if (!hasToolCall) {
              this.continuationNeeded = true;
              this.continuationReason = "length";
            }
          }
          const isError =
            overflow ||
            assistantMsg.stopReason === "error" ||
            assistantMsg.stopReason === "aborted";
          const streamId = this.streamingMessageId;

          this.updateMessages(
            (msgs) => {
              const messages = [...msgs];
              const idx = messages.findIndex((m) => m.id === streamId);

              if (isError) {
                if (idx !== -1) {
                  messages.splice(idx, 1);
                }
              } else if (idx !== -1) {
                const parts = extractPartsFromAssistantMessage(
                  event.message,
                  messages[idx].parts,
                );
                messages[idx] = { ...messages[idx], parts };
              }
              return messages;
            },
            {
              error: isError
                ? assistantMsg.errorMessage || "Request failed"
                : this.state.error,
              sessionStats: isError
                ? this.state.sessionStats
                : {
                    ...deriveStats(this.agent?.state.messages ?? []),
                    contextWindow: this.state.sessionStats.contextWindow,
                  },
            },
          );
          this.streamingMessageId = null;
        }
        break;
      }
      case "tool_execution_start": {
        this.recordToolStart(event.toolCallId, event.toolName, event.args);
        this.updateMessages((msgs) => {
          const messages = [...msgs];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex(
              (p) => p.type === "toolCall" && p.id === event.toolCallId,
            );
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                parts[partIdx] = { ...part, status: "running" };
                messages[i] = { ...msg, parts };
              }
              break;
            }
          }
          return messages;
        });
        break;
      }
      case "tool_execution_update": {
        this.updateMessages((msgs) => {
          const messages = [...msgs];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex(
              (p) => p.type === "toolCall" && p.id === event.toolCallId,
            );
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                let partialText: string;
                if (typeof event.partialResult === "string") {
                  partialText = event.partialResult;
                } else if (
                  event.partialResult?.content &&
                  Array.isArray(event.partialResult.content)
                ) {
                  partialText = event.partialResult.content
                    .filter((c: { type: string }) => c.type === "text")
                    .map((c: { text: string }) => c.text)
                    .join("\n");
                } else {
                  partialText = JSON.stringify(event.partialResult, null, 2);
                }
                parts[partIdx] = { ...part, result: partialText };
                messages[i] = { ...msg, parts };
              }
              break;
            }
          }
          return messages;
        });
        break;
      }
      case "tool_execution_end": {
        const toolFailed = toolResultFailed(event.result, event.isError);
        this.recordToolEnd(
          event.toolCallId,
          event.toolName,
          undefined,
          event.result,
          event.isError,
        );
        let resultText: string;
        let resultImages: { data: string; mimeType: string }[] | undefined;
        if (typeof event.result === "string") {
          resultText = event.result;
        } else if (
          event.result?.content &&
          Array.isArray(event.result.content)
        ) {
          resultText = event.result.content
            .filter((c: { type: string }) => c.type === "text")
            .map((c: { text: string }) => c.text)
            .join("\n");
          const images = event.result.content
            .filter((c: { type: string }) => c.type === "image")
            .map((c: { data: string; mimeType: string }) => ({
              data: c.data,
              mimeType: c.mimeType,
            }));
          if (images.length > 0) resultImages = images;
        } else {
          resultText = JSON.stringify(event.result, null, 2);
        }

        if (!toolFailed && this.followMode) {
          this.adapter.onToolResult?.(event.toolCallId, resultText, false);
        }

        this.updateMessages((msgs) => {
          const messages = [...msgs];
          for (let i = messages.length - 1; i >= 0; i--) {
            const msg = messages[i];
            const partIdx = msg.parts.findIndex(
              (p) => p.type === "toolCall" && p.id === event.toolCallId,
            );
            if (partIdx !== -1) {
              const parts = [...msg.parts];
              const part = parts[partIdx];
              if (part.type === "toolCall") {
                parts[partIdx] = {
                  ...part,
                  status: toolFailed ? "error" : "complete",
                  result: resultText,
                  images: resultImages,
                };
                messages[i] = { ...msg, parts };
              }
              break;
            }
          }
          return messages;
        });
        break;
      }
      case "agent_end": {
        if (this.manualAbortPending) {
          this.captureRecoveryMutations();
          this.manualAbortPending = false;
        }
        if (this.continuationNeeded && this.continuationReason === "overflow") {
          this.streamingMessageId = null;
          break;
        }
        this.streamingMessageId = null;
        await this.onStreamingEnd();
        this.isStreaming = false;
        this.update({ isStreaming: false });
        break;
      }
    }
  };

  applyConfig(config: ProviderConfig) {
    let contextWindow = 0;
    let baseModel: Model<Api>;
    if (config.provider === "custom") {
      const custom = buildCustomModel(config);
      if (!custom) return;
      baseModel = custom;
    } else {
      try {
        baseModel = (getModel as (p: string, m: string) => Model<Api>)(
          config.provider,
          config.model,
        );
      } catch {
        return;
      }
    }
    contextWindow = baseModel.contextWindow;
    this.config = config;

    const proxiedModel = applyProxyToModel(baseModel, config);
    const existingMessages = this.agent?.state.messages ?? [];

    if (this.agent) {
      this.agent.abort();
    }

    const systemPrompt = this.adapter.buildSystemPrompt(
      this.skills,
      this.context.commandSnippets,
    );

    const agent = new Agent({
      initialState: {
        model: proxiedModel,
        systemPrompt,
        thinkingLevel: thinkingLevelToAgent(config.thinking),
        tools: this.tools,
        messages: existingMessages,
      },
      convertToLlm: (messages): Message[] =>
        messages.flatMap((message) => {
          if (
            message.role !== "user" &&
            message.role !== "assistant" &&
            message.role !== "toolResult"
          ) {
            return [];
          }
          const llmMessage = { ...message } as Message & {
            _runtimeUsage?: RuntimeUsageCarry;
          };
          delete llmMessage._runtimeUsage;
          return [llmMessage];
        }),
      streamFn: async (model, context, options) => {
        const cfg = this.config ?? config;
        const apiKey = await this.getActiveApiKey(cfg);
        return streamSimple(model, context, {
          ...options,
          apiKey,
        });
      },
      transformContext: this.adapter.transformContext
        ? (messages, signal) =>
            this.adapter.transformContext!(messages, signal, {
              contextWindow,
              systemPromptChars:
                this.agent?.state.systemPrompt.length ?? systemPrompt.length,
              recoveryAttempt: this.contextRecoveryAttempt,
            })
        : undefined,
      toolExecution: this.adapter.toolExecution,
      beforeToolCall: async ({ toolCall, args }, signal) => {
        if (signal?.aborted) {
          return {
            block: true,
            reason:
              "RUNTIME_ABORTED: tool execution was cancelled before it started.",
          };
        }
        const info = this.classifyTool(toolCall.name, args);
        if (!info) return undefined;
        if (this.isReplayGuarded(toolCall.name, args)) {
          return {
            block: true,
            reason:
              "RUNTIME_REPLAY_GUARD: this completed document write was already applied before interruption. Inspect the current document and apply only missing changes.",
          };
        }
        if (this.recoveryVerificationRequired && info.effect !== "read") {
          return {
            block: true,
            reason:
              "RUNTIME_RECOVERY_VERIFY_FIRST: inspect the current document with a targeted read or verification tool before any new write.",
          };
        }
        if (signal?.aborted) {
          return {
            block: true,
            reason:
              "RUNTIME_ABORTED: tool execution was cancelled before it started.",
          };
        }
        this.markToolAdmitted(toolCall.id);
        return undefined;
      },
    });
    this.agent = agent;
    agent.subscribe(this.handleAgentEvent);
    this.pendingConfig = null;
    this.followMode = config.followMode ?? true;

    this.update({
      providerConfig: config,
      error: null,
      sessionStats: {
        ...this.state.sessionStats,
        contextWindow,
      },
    });
  }

  setProviderConfig(config: ProviderConfig) {
    if (this.isStreaming || this.agent?.state.isStreaming) {
      this.pendingConfig = config;
      this.update({ providerConfig: config });
      return;
    }
    this.applyConfig(config);
  }

  abort() {
    this.sendGeneration++;
    this.pendingSendController?.abort();
    this.pendingSendController = null;
    this.lastRunSucceeded = false;
    if (this.agent?.state.isStreaming) {
      this.manualAbortPending = true;
      this.captureRecoveryMutations();
    }
    this.agent?.abort();
    this.agent?.clearAllQueues();
    this.continuationNeeded = false;
    this.continuationReason = null;
    this.autoContinuationCount = 0;
    this.contextRecoveryAttempt = 0;
    this.isStreaming = false;
    this.update({ isStreaming: false });
  }

  async sendMessage(content: string, attachments?: string[]) {
    this.pendingSendController?.abort();
    const sendController = new AbortController();
    const sendGeneration = ++this.sendGeneration;
    this.pendingSendController = sendController;
    let pendingUserMessageId: string | null = null;
    let promptStarted = false;
    const isCancelled = () =>
      sendController.signal.aborted || sendGeneration !== this.sendGeneration;
    const finishCancelled = () => {
      if (!isCancelled()) return false;
      if (!promptStarted && pendingUserMessageId) {
        const messageId = pendingUserMessageId;
        pendingUserMessageId = null;
        this.updateMessages((messages) =>
          messages.filter((message) => message.id !== messageId),
        );
      }
      if (this.pendingSendController === sendController) {
        this.pendingSendController = null;
      }
      if (sendGeneration === this.sendGeneration) {
        this.isStreaming = false;
        this.update({ isStreaming: false });
      }
      return true;
    };
    const releaseSend = () => {
      if (this.pendingSendController === sendController) {
        this.pendingSendController = null;
      }
    };

    let agent = this.agent;
    if (agent?.state.isStreaming) {
      await agent.waitForIdle();
      if (finishCancelled()) return;
    }
    if (this.pendingConfig) {
      this.applyConfig(this.pendingConfig);
      agent = this.agent;
    }
    if (!agent || !this.state.providerConfig) {
      releaseSend();
      this.update({ error: "Please configure your API key first" });
      return;
    }
    if (agent.state.isStreaming) {
      await agent.waitForIdle();
      if (finishCancelled()) return;
    }
    if (this.manualAbortPending) {
      this.captureRecoveryMutations();
      this.manualAbortPending = false;
    }
    this.toolJournal.clear();
    this.verifiedRecoveryIds.clear();

    const userMessage: ChatMessage = {
      id: generateId(),
      role: "user",
      parts: [{ type: "text", text: content }],
      timestamp: Date.now(),
    };
    pendingUserMessageId = userMessage.id;

    this.isStreaming = true;
    this.update({
      messages: [...this.state.messages, userMessage],
      isStreaming: true,
      error: null,
    });

    try {
      let promptContent = content;
      this.continuationNeeded = false;
      this.continuationReason = null;
      this.autoContinuationCount = 0;
      this.contextRecoveryAttempt = 0;
      agent.clearAllQueues();

      const messageInfo = this.getMessagePreparationInfo(agent);
      const messageTools = this.adapter.toolsForMessage
        ? this.adapter.toolsForMessage(content, this.context, messageInfo)
        : this.tools;
      agent.state.tools = messageTools;

      let messageSystemPrompt: string | undefined;
      if (this.adapter.buildSystemPromptForMessage) {
        messageSystemPrompt = this.adapter.buildSystemPromptForMessage(
          content,
          this.skills,
          this.context.commandSnippets,
          messageInfo,
        );
        agent.state.systemPrompt = messageSystemPrompt;
      }

      if (this.adapter.getDocumentMetadata) {
        try {
          if (this.adapter.metadataHistory === "latest") {
            this.stripHistoricalMetadata();
          }
          const meta = await this.adapter.getDocumentMetadata({
            userMessage: content,
            messageCount: agent.state.messages.length,
            signal: sendController.signal,
            info: messageInfo,
          });
          if (finishCancelled()) return;
          if (meta) {
            const tag = this.adapter.metadataTag || "doc_context";
            promptContent = `<${tag}>\n${JSON.stringify(meta.metadata, null, 2)}\n</${tag}>\n\n${content}`;
            if (meta.nameMap) {
              this.update({ nameMap: meta.nameMap });
            }
          }
        } catch (err) {
          if (finishCancelled()) return;
          console.error("[Runtime] Failed to get document metadata:", err);
        }
      }

      // A provider change while document metadata is loading is queued because
      // the runtime is busy even though the old agent has not started prompting.
      // Apply it at the last pre-prompt boundary and retain this turn's routing.
      if (this.pendingConfig) {
        this.applyConfig(this.pendingConfig);
        agent = this.agent ?? agent;
        agent.state.tools = messageTools;
        if (messageSystemPrompt !== undefined) {
          agent.state.systemPrompt = messageSystemPrompt;
        }
      }

      if (attachments && attachments.length > 0) {
        const paths = attachments
          .map((name) => `/home/user/uploads/${name}`)
          .join("\n");
        promptContent = `<attachments>\n${paths}\n</attachments>\n\n${promptContent}`;
      }

      if (this.pendingRecovery.size > 0) {
        promptContent = `${this.buildRecoveryPrompt(
          "resume_after_abort",
          1,
        )}\n\n${promptContent}`;
      }

      const promptMessageIndex = agent.state.messages.length;
      if (finishCancelled()) return;
      this.lastRunSucceeded = false;
      promptStarted = true;
      await agent.prompt(promptContent);
      if (finishCancelled()) return;

      // Read-only overflow can safely replay the turn after compaction. If a
      // document mutation may already have landed, preserve its tool receipt
      // and continue from a hidden inspect-before-write recovery prompt.
      let continuationAttempts = 0;
      while (
        this.continuationNeeded &&
        this.continuationReason === "overflow" &&
        !!this.adapter.transformContext &&
        continuationAttempts < 2
      ) {
        continuationAttempts++;
        this.contextRecoveryAttempt = continuationAttempts;
        this.continuationNeeded = false;
        this.lastRunSucceeded = false;
        this.isStreaming = true;
        this.update({ isStreaming: true, error: null });
        if (
          this.currentRecoveryMutations().length > 0 ||
          this.pendingRecovery.size > 0
        ) {
          const recoveryPrompt = this.prepareRecovery(
            promptMessageIndex,
            "overflow",
            continuationAttempts,
          );
          await agent.prompt(recoveryPrompt);
        } else {
          this.rollbackOverflowTurn(promptMessageIndex);
          await agent.continue();
        }
        if (finishCancelled()) return;
      }
      if (!this.continuationNeeded && this.lastRunSucceeded) {
        this.settleResolvedRecovery();
      }
      if (this.continuationNeeded) {
        const reason = this.continuationReason;
        const error =
          reason === "overflow"
            ? "上下文已接近模型上限，已自动压缩但仍未完成。请缩小处理范围后继续。"
            : "模型单次回复达到上限。请发送“继续”，当前对话仍保留。";
        if (reason === "overflow") {
          if (
            this.currentRecoveryMutations().length > 0 ||
            this.pendingRecovery.size > 0
          ) {
            this.prepareRecovery(
              promptMessageIndex,
              "overflow",
              continuationAttempts + 1,
            );
          } else {
            this.rollbackOverflowTurn(promptMessageIndex);
          }
          this.update({
            error,
            sessionStats: {
              ...deriveStats(agent.state.messages),
              contextWindow: this.state.sessionStats.contextWindow,
            },
          });
          await this.onStreamingEnd();
          this.isStreaming = false;
          this.update({ isStreaming: false });
        } else {
          this.update({ error });
        }
      }
    } catch (err) {
      if (finishCancelled()) return;
      console.error("[Runtime] sendMessage error:", err);
      this.isStreaming = false;
      this.update({
        isStreaming: false,
        error: err instanceof Error ? err.message : "An error occurred",
      });
    }
    releaseSend();
  }

  private rollbackOverflowTurn(promptMessageIndex: number): void {
    const messages = this.agent?.state.messages;
    if (!messages || promptMessageIndex < 0) return;
    const prompt = messages[promptMessageIndex];
    if (prompt?.role === "user") {
      const carrier = prompt as typeof prompt & {
        _runtimeUsage?: RuntimeUsageCarry;
      };
      const carry: RuntimeUsageCarry = carrier._runtimeUsage ?? {
        inputTokens: 0,
        outputTokens: 0,
        cacheRead: 0,
        cacheWrite: 0,
        totalCost: 0,
        lastInputTokens: 0,
      };
      for (const message of messages.slice(promptMessageIndex + 1)) {
        if (message.role !== "assistant") continue;
        const usage = (message as AssistantMessage).usage;
        if (!usage) continue;
        carry.inputTokens += usage.input;
        carry.outputTokens += usage.output;
        carry.cacheRead += usage.cacheRead;
        carry.cacheWrite += usage.cacheWrite;
        carry.totalCost += usage.cost.total;
        carry.lastInputTokens =
          usage.input + usage.cacheRead + usage.cacheWrite;
      }
      carrier._runtimeUsage = carry;
    }
    const keepCount =
      prompt?.role === "user" ? promptMessageIndex + 1 : promptMessageIndex;
    messages.splice(keepCount);
  }

  /** Remove stale per-turn document snapshots while keeping the latest one. */
  private stripHistoricalMetadata(): void {
    const tag = this.adapter.metadataTag || "doc_context";
    const escaped = tag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const pattern = new RegExp(
      `<${escaped}>\\r?\\n[\\s\\S]*?\\r?\\n</${escaped}>\\r?\\n\\r?\\n`,
    );

    const messages = this.agent?.state.messages ?? [];
    for (const message of messages) {
      if (message.role !== "user") continue;
      if (typeof message.content === "string") {
        message.content = message.content.replace(pattern, "");
      } else if (Array.isArray(message.content)) {
        message.content = message.content.map((block) =>
          block.type === "text"
            ? { ...block, text: block.text.replace(pattern, "") }
            : block,
        );
      }
    }
  }

  async clearMessages() {
    this.abort();
    const generation = this.sendGeneration;
    const agent = this.agent;
    await agent?.waitForIdle();
    if (generation !== this.sendGeneration || agent !== this.agent) return;
    this.continuationNeeded = false;
    this.continuationReason = null;
    this.autoContinuationCount = 0;
    this.contextRecoveryAttempt = 0;
    this.resetRecoveryState();
    this.agent?.reset();
    this.context.reset();
    if (this.currentSessionId) {
      Promise.all([
        saveSession(this.ns, this.currentSessionId, []),
        saveVfsFiles(this.ns, this.currentSessionId, []),
      ]).catch(console.error);
    }
    this.update({
      messages: [],
      error: null,
      sessionStats: INITIAL_STATS,
      uploads: [],
    });
  }

  private async refreshSessions() {
    if (!this.documentId) return;
    const sessions = await listSessions(this.ns, this.documentId);
    this.update({ sessions });
  }

  async newSession() {
    if (!this.documentId) return;
    if (this.isStreaming) return;
    const generation = ++this.sendGeneration;
    this.pendingSendController?.abort();
    this.pendingSendController = null;
    const agent = this.agent;
    await agent?.waitForIdle();
    if (
      generation !== this.sendGeneration ||
      agent !== this.agent ||
      this.isStreaming
    ) {
      return;
    }
    try {
      this.continuationNeeded = false;
      this.continuationReason = null;
      this.autoContinuationCount = 0;
      this.contextRecoveryAttempt = 0;
      this.resetRecoveryState();
      this.agent?.reset();
      this.context.reset();
      const session = await createSession(this.ns, this.documentId);
      this.currentSessionId = session.id;
      await this.refreshSessions();
      this.update({
        messages: [],
        currentSession: session,
        error: null,
        sessionStats: INITIAL_STATS,
        uploads: [],
      });
    } catch (err) {
      console.error("[Runtime] Failed to create session:", err);
    }
  }

  async switchSession(sessionId: string) {
    if (this.currentSessionId === sessionId) return;
    if (this.isStreaming) return;
    const generation = ++this.sendGeneration;
    this.pendingSendController?.abort();
    this.pendingSendController = null;
    const agent = this.agent;
    await agent?.waitForIdle();
    if (
      generation !== this.sendGeneration ||
      agent !== this.agent ||
      this.isStreaming
    ) {
      return;
    }
    this.continuationNeeded = false;
    this.continuationReason = null;
    this.autoContinuationCount = 0;
    this.contextRecoveryAttempt = 0;
    this.resetRecoveryState();
    this.agent?.reset();
    try {
      const [session, vfsFiles] = await Promise.all([
        getSession(this.ns, sessionId),
        loadVfsFiles(this.ns, sessionId),
      ]);
      if (!session) return;
      await this.context.restoreVfs(vfsFiles);
      this.currentSessionId = session.id;

      if (session.agentMessages.length > 0 && this.agent) {
        this.agent.state.messages = session.agentMessages;
      }

      const uploadNames = await this.context.listUploads();
      const stats = deriveStats(session.agentMessages);
      this.update({
        messages: agentMessagesToChatMessages(
          session.agentMessages,
          this.adapter.metadataTag,
        ),
        currentSession: session,
        error: null,
        sessionStats: {
          ...stats,
          contextWindow: this.state.sessionStats.contextWindow,
        },
        uploads: uploadNames.map((name) => ({ name, size: 0 })),
      });
      await this.refreshNameMap();
    } catch (err) {
      console.error("[Runtime] Failed to switch session:", err);
    }
  }

  async deleteCurrentSession() {
    if (!this.currentSessionId || !this.documentId) return;
    if (this.isStreaming) return;
    const generation = ++this.sendGeneration;
    this.pendingSendController?.abort();
    this.pendingSendController = null;
    const agent = this.agent;
    await agent?.waitForIdle();
    if (
      generation !== this.sendGeneration ||
      agent !== this.agent ||
      this.isStreaming
    ) {
      return;
    }
    this.resetRecoveryState();
    this.agent?.reset();
    const deletedId = this.currentSessionId;
    await Promise.all([
      deleteSession(this.ns, deletedId),
      saveVfsFiles(this.ns, deletedId, []),
    ]);
    const session = await getOrCreateCurrentSession(this.ns, this.documentId);
    this.currentSessionId = session.id;
    const vfsFiles = await loadVfsFiles(this.ns, session.id);
    await this.context.restoreVfs(vfsFiles);

    if (session.agentMessages.length > 0 && this.agent) {
      this.agent.state.messages = session.agentMessages;
    }

    await this.refreshSessions();
    const uploadNames = await this.context.listUploads();
    const stats = deriveStats(session.agentMessages);
    this.update({
      messages: agentMessagesToChatMessages(
        session.agentMessages,
        this.adapter.metadataTag,
      ),
      currentSession: session,
      error: null,
      sessionStats: {
        ...stats,
        contextWindow: this.state.sessionStats.contextWindow,
      },
      uploads: uploadNames.map((name) => ({ name, size: 0 })),
    });
  }

  private async onStreamingEnd() {
    if (!this.currentSessionId) return;
    const sessionId = this.currentSessionId;
    const agentMessages = this.agent?.state.messages ?? [];
    try {
      const vfsFiles = await this.context.snapshotVfs();
      await Promise.all([
        saveSession(this.ns, sessionId, agentMessages),
        saveVfsFiles(this.ns, sessionId, vfsFiles),
      ]);
      await this.refreshSessions();
      const updated = await getSession(this.ns, sessionId);
      if (updated && this.currentSessionId === sessionId) {
        this.update({ currentSession: updated });
      }
      if (this.currentSessionId === sessionId) this.bumpVfs();
    } catch (e) {
      console.error(e);
    }
  }

  async init() {
    if (this.sessionLoaded) return;
    this.sessionLoaded = true;

    try {
      // update staticFiles and customCommands jic any changes
      // happened between context init (on app mount) vs session init
      if (this.adapter.staticFiles) {
        await this.context.setStaticFiles(this.adapter.staticFiles);
      }
      if (this.adapter.customCommands) {
        this.context.setCustomCommands(this.adapter.customCommands);
      }

      const id = await this.adapter.getDocumentId();
      this.documentId = id;

      const skills = await getInstalledSkills(this.ns);
      this.skills = skills;
      await syncSkillsToVfs(this.ns, this.context);

      const saved = loadSavedConfig(this.ns);
      if (saved?.provider && saved?.apiKey && saved?.model) {
        this.applyConfig(saved);
      }

      const session = await getOrCreateCurrentSession(this.ns, id);
      this.currentSessionId = session.id;
      const [sessions, vfsFiles] = await Promise.all([
        listSessions(this.ns, id),
        loadVfsFiles(this.ns, session.id),
      ]);
      if (vfsFiles.length > 0) {
        await this.context.restoreVfs(vfsFiles);
      }

      if (session.agentMessages.length > 0 && this.agent) {
        this.agent.state.messages = session.agentMessages;
      }

      const uploadNames = await this.context.listUploads();
      const stats = deriveStats(session.agentMessages);
      this.update({
        messages: agentMessagesToChatMessages(
          session.agentMessages,
          this.adapter.metadataTag,
        ),
        currentSession: session,
        sessions,
        skills,
        sessionStats: {
          ...stats,
          contextWindow: this.state.sessionStats.contextWindow,
        },
        uploads: uploadNames.map((name) => ({ name, size: 0 })),
      });
      await this.refreshNameMap();
    } catch (err) {
      console.error("[Runtime] Failed to load session:", err);
    }
  }

  async uploadFiles(files: { name: string; size: number; data: Uint8Array }[]) {
    if (files.length === 0) return;
    this.uploadOperationCount++;
    this.update({ isUploading: true, error: null });
    const operation = this.uploadQueue
      .catch(() => {})
      .then(async () => {
        for (const file of files) {
          await this.context.writeFile(file.name, file.data);
          const uploads = [...this.state.uploads];
          const exists = uploads.findIndex((u) => u.name === file.name);
          if (exists !== -1) {
            uploads[exists] = { name: file.name, size: file.size };
          } else {
            uploads.push({ name: file.name, size: file.size });
          }
          this.update({ uploads });
        }
        if (this.currentSessionId) {
          const snapshot = await this.context.snapshotVfs();
          await saveVfsFiles(this.ns, this.currentSessionId, snapshot);
        }
        this.bumpVfs();
      });
    this.uploadQueue = operation.catch(() => {});
    try {
      await operation;
    } catch (err) {
      console.error("Failed to upload file:", err);
      this.update({
        error:
          err instanceof Error
            ? `Failed to upload file: ${err.message}`
            : "Failed to upload file",
      });
    } finally {
      this.uploadOperationCount = Math.max(0, this.uploadOperationCount - 1);
      if (this.uploadOperationCount === 0) {
        this.update({ isUploading: false });
      }
    }
  }

  async removeUpload(name: string) {
    try {
      await this.context.deleteFile(name);
      this.update({
        uploads: this.state.uploads.filter((u) => u.name !== name),
      });
      if (this.currentSessionId) {
        const snapshot = await this.context.snapshotVfs();
        await saveVfsFiles(this.ns, this.currentSessionId, snapshot);
      }
      this.bumpVfs();
    } catch (err) {
      console.error("Failed to delete file:", err);
      this.update({
        uploads: this.state.uploads.filter((u) => u.name !== name),
      });
    }
  }

  private async refreshSkillsAndRebuildAgent() {
    this.skills = await getInstalledSkills(this.ns);
    this.update({ skills: this.skills });
    if (this.state.providerConfig) {
      this.applyConfig(this.state.providerConfig);
    }
  }

  async installSkill(inputs: { path: string; data: Uint8Array }[]) {
    if (inputs.length === 0) return;
    try {
      await addSkill(this.ns, this.context, inputs);
      await this.refreshSkillsAndRebuildAgent();
    } catch (err) {
      console.error("[Runtime] Failed to install skill:", err);
      this.update({
        error: err instanceof Error ? err.message : "Failed to install skill",
      });
    }
  }

  async uninstallSkill(name: string) {
    try {
      await removeSkill(this.ns, this.context, name);
      await this.refreshSkillsAndRebuildAgent();
    } catch (err) {
      console.error("[Runtime] Failed to uninstall skill:", err);
    }
  }

  toggleFollowMode() {
    if (!this.state.providerConfig) return;
    const newFollowMode = !this.state.providerConfig.followMode;
    this.followMode = newFollowMode;
    const newConfig = {
      ...this.state.providerConfig,
      followMode: newFollowMode,
    };
    saveConfig(this.ns, newConfig);
    this.update({ providerConfig: newConfig });
  }

  toggleExpandToolCalls() {
    if (!this.state.providerConfig) return;
    const newConfig = {
      ...this.state.providerConfig,
      expandToolCalls: !this.state.providerConfig.expandToolCalls,
    };
    saveConfig(this.ns, newConfig);
    this.update({ providerConfig: newConfig });
  }

  getName(id: number): string | undefined {
    return this.state.nameMap[id];
  }

  private async refreshNameMap() {
    if (!this.adapter.getDocumentMetadata) return;
    try {
      const meta = await this.adapter.getDocumentMetadata();
      if (meta?.nameMap) {
        this.update({ nameMap: meta.nameMap });
      }
    } catch (err) {
      console.error("[Runtime] Failed to refresh nameMap:", err);
    }
  }

  dispose() {
    this.agent?.abort();
    this.listeners.clear();
  }
}
