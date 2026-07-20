import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type {
  AssistantMessage,
  ImageContent,
  TextContent,
  ToolResultMessage,
  UserMessage,
} from "@earendil-works/pi-ai";

export type ToolCallStatus = "pending" | "running" | "complete" | "error";

export type MessagePart =
  | { type: "text"; text: string }
  | { type: "thinking"; thinking: string }
  | {
      type: "toolCall";
      id: string;
      name: string;
      args: Record<string, unknown>;
      status: ToolCallStatus;
      result?: string;
      images?: { data: string; mimeType: string }[];
    };

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  parts: MessagePart[];
  timestamp: number;
}

export interface SessionStats {
  inputTokens: number;
  outputTokens: number;
  cacheRead: number;
  cacheWrite: number;
  totalCost: number;
  contextWindow: number;
  lastInputTokens: number;
}

export interface RuntimeUsageCarry {
  inputTokens: number;
  outputTokens: number;
  cacheRead: number;
  cacheWrite: number;
  totalCost: number;
  lastInputTokens: number;
}

function contentToText(
  content: string | { type: string; text?: string }[],
): string {
  if (typeof content === "string") return content;
  return content
    .filter((c) => c.type === "text")
    .map((c) => c.text ?? "")
    .join("\n");
}

function isRuntimeControlMessage(
  content: string | { type: string; text?: string }[],
): boolean {
  const text = contentToText(content).trim();
  return (
    /^<runtime_continue>[\s\S]*<\/runtime_continue>$/.test(text) ||
    /^<runtime_recovery>[\s\S]*<\/runtime_recovery>$/.test(text)
  );
}

export function stripEnrichment(
  content: string | { type: string; text?: string }[],
  metadataTag?: string,
): string {
  let text = contentToText(content);
  text = text.replace(/^<runtime_continue>[\s\S]*<\/runtime_continue>$/, "");
  text = text.replace(
    /^<runtime_recovery>[\s\S]*?<\/runtime_recovery>\r?\n\r?\n/,
    "",
  );
  text = text.replace(/^<runtime_recovery>[\s\S]*<\/runtime_recovery>$/, "");
  text = text.replace(/^<attachments>\n[\s\S]*?\n<\/attachments>\n\n/, "");
  if (metadataTag) {
    const escaped = metadataTag.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    text = text.replace(
      new RegExp(`^<${escaped}>\\n[\\s\\S]*?\\n</${escaped}>\\n\\n`),
      "",
    );
  } else {
    text = text.replace(/^<\w+_context>\n[\s\S]*?\n<\/\w+_context>\n\n/, "");
  }
  return text;
}

export function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

export function extractPartsFromAssistantMessage(
  message: AgentMessage,
  existingParts: MessagePart[] = [],
): MessagePart[] {
  if (message.role !== "assistant") return existingParts;

  const assistantMsg = message as AssistantMessage;
  const existingToolCalls = new Map<string, MessagePart>();
  for (const part of existingParts) {
    if (part.type === "toolCall") {
      existingToolCalls.set(part.id, part);
    }
  }

  return assistantMsg.content.map((block): MessagePart => {
    if (block.type === "text") {
      return { type: "text", text: block.text };
    }
    if (block.type === "thinking") {
      return { type: "thinking", thinking: block.thinking };
    }
    const existing = existingToolCalls.get(block.id);
    return {
      type: "toolCall",
      id: block.id,
      name: block.name,
      args: block.arguments as Record<string, unknown>,
      status: existing?.type === "toolCall" ? existing.status : "pending",
      result: existing?.type === "toolCall" ? existing.result : undefined,
    };
  });
}

export function agentMessagesToChatMessages(
  agentMessages: AgentMessage[],
  metadataTag?: string,
): ChatMessage[] {
  const result: ChatMessage[] = [];
  for (const msg of agentMessages) {
    if (msg.role === "user") {
      if (isRuntimeControlMessage((msg as UserMessage).content)) continue;
      const text = stripEnrichment((msg as UserMessage).content, metadataTag);
      result.push({
        id: generateId(),
        role: "user",
        parts: [{ type: "text", text }],
        timestamp: msg.timestamp,
      });
    } else if (msg.role === "assistant") {
      const parts = extractPartsFromAssistantMessage(msg);
      result.push({
        id: generateId(),
        role: "assistant",
        parts,
        timestamp: msg.timestamp,
      });
    } else if (msg.role === "toolResult") {
      const toolResult = msg as ToolResultMessage;
      for (let i = result.length - 1; i >= 0; i--) {
        const chatMsg = result[i];
        if (chatMsg.role !== "assistant") continue;
        const partIdx = chatMsg.parts.findIndex(
          (p) => p.type === "toolCall" && p.id === toolResult.toolCallId,
        );
        if (partIdx !== -1) {
          const part = chatMsg.parts[partIdx];
          if (part.type === "toolCall") {
            const resultText = toolResult.content
              .filter((c): c is TextContent => c.type === "text")
              .map((c) => c.text)
              .join("\n");
            const images = toolResult.content
              .filter((c): c is ImageContent => c.type === "image")
              .map((c) => ({ data: c.data, mimeType: c.mimeType }));
            chatMsg.parts[partIdx] = {
              ...part,
              status: toolResult.isError ? "error" : "complete",
              result: resultText,
              images: images.length > 0 ? images : undefined,
            };
          }
          break;
        }
      }
    }
  }
  return result;
}

export function deriveStats(
  agentMessages: AgentMessage[],
): Omit<SessionStats, "contextWindow"> {
  let inputTokens = 0;
  let outputTokens = 0;
  let cacheRead = 0;
  let cacheWrite = 0;
  let totalCost = 0;
  let lastInputTokens = 0;
  for (const msg of agentMessages) {
    const carry = (msg as AgentMessage & { _runtimeUsage?: RuntimeUsageCarry })
      ._runtimeUsage;
    if (carry) {
      inputTokens += carry.inputTokens;
      outputTokens += carry.outputTokens;
      cacheRead += carry.cacheRead;
      cacheWrite += carry.cacheWrite;
      totalCost += carry.totalCost;
      lastInputTokens = carry.lastInputTokens;
    }
    if (msg.role === "assistant") {
      const u = (msg as AssistantMessage).usage;
      if (u) {
        inputTokens += u.input;
        outputTokens += u.output;
        cacheRead += u.cacheRead;
        cacheWrite += u.cacheWrite;
        totalCost += u.cost.total;
        lastInputTokens = u.input + u.cacheRead + u.cacheWrite;
      }
    }
  }
  return {
    inputTokens,
    outputTokens,
    cacheRead,
    cacheWrite,
    totalCost,
    lastInputTokens,
  };
}
