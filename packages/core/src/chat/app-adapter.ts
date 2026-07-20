import type {
  AgentTool,
  ToolExecutionMode,
} from "@earendil-works/pi-agent-core";
import type {
  AgentContext,
  CustomCommandsResult,
  MessagePreparationInfo,
  RecoveryPromptInput,
  SkillMeta,
  StorageNamespace,
  ToolRecoveryInfo,
} from "@office-agents/sdk";

export type { MessagePreparationInfo, StorageNamespace };

import type { Component } from "svelte";

export type { CustomCommandsResult };

export type MaybePromise<T> = T | Promise<T>;

/** Context passed to an app adapter before a user message is sent. */
export interface DocumentMetadataRequest {
  userMessage: string;
  messageCount: number;
  signal?: AbortSignal;
  info?: MessagePreparationInfo;
}

export type DocumentMetadata = {
  metadata: object;
  nameMap?: Record<number, string>;
};

export interface ContextTransformInfo {
  contextWindow: number;
  systemPromptChars: number;
  recoveryAttempt: number;
}

export interface LinkClickContext {
  href: string;
  anchor: HTMLAnchorElement;
  event: MouseEvent;
}

export type LinkClickResult = "handled" | "default";

export interface ToolExtrasProps {
  toolName: string;
  result?: string;
  expanded: boolean;
}

export interface AppAdapter {
  tools: AgentTool[] | ((ctx: AgentContext) => AgentTool[]);
  /** Optional task-aware tool allowlist rebuilt before each user message. */
  toolsForMessage?: (
    userMessage: string,
    ctx: AgentContext,
    info: MessagePreparationInfo,
  ) => AgentTool[];
  buildSystemPrompt: (skills: SkillMeta[], commandSnippets: string[]) => string;
  /** Optional task-aware system prompt. Called immediately before each prompt. */
  buildSystemPromptForMessage?: (
    userMessage: string,
    skills: SkillMeta[],
    commandSnippets: string[],
    info?: MessagePreparationInfo,
  ) => string;
  getDocumentId: () => Promise<string>;
  getDocumentMetadata?: (
    request?: DocumentMetadataRequest,
  ) => Promise<DocumentMetadata | null>;
  /**
   * Keep only the latest metadata block in the agent transcript. This is
   * useful for documents whose selection changes between turns.
   */
  metadataHistory?: "all" | "latest";
  /** Classify document effects so interrupted writes are never blindly replayed. */
  getToolRecoveryInfo?: (
    toolName: string,
    args: unknown,
  ) => ToolRecoveryInfo | undefined;
  normalizeToolArgsForReplay?: (toolName: string, args: unknown) => unknown;
  buildRecoveryPrompt?: (input: RecoveryPromptInput) => string;
  toolExecution?: ToolExecutionMode;
  /** Optional app-specific context pruning before an LLM request. */
  transformContext?: (
    messages: import("@earendil-works/pi-agent-core").AgentMessage[],
    signal?: AbortSignal,
    info?: ContextTransformInfo,
  ) => Promise<import("@earendil-works/pi-agent-core").AgentMessage[]>;
  onToolResult?: (toolCallId: string, result: string, isError: boolean) => void;
  metadataTag?: string;
  storageNamespace?: Partial<StorageNamespace>;
  appVersion?: string;
  appName?: string;
  emptyStateMessage?: string;
  staticFiles?: Record<string, string>;
  customCommands?: (ns: StorageNamespace) => CustomCommandsResult;
  hasImageSearch?: boolean;
  showFollowModeToggle?: boolean;
  handleLinkClick?: (
    context: LinkClickContext,
  ) => MaybePromise<LinkClickResult>;
  ToolExtras?: Component<ToolExtrasProps>;
  HeaderExtras?: Component;
  SelectionIndicator?: Component;
}
