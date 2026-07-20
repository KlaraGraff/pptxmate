import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { resizeImage } from "@office-agents/core";
import type { Static, TObject } from "@sinclair/typebox";
import { capText, TOOL_RESULT_MAX_BYTES } from "./result-budget";

export type ToolResult = AgentToolResult<undefined>;

interface ToolConfig<T extends TObject> {
  name: string;
  label: string;
  description: string;
  parameters: T;
  execute: (
    toolCallId: string,
    params: Static<T>,
    signal?: AbortSignal,
  ) => Promise<ToolResult>;
  modifiedSlide?: (
    params: Static<T>,
    result?: Record<string, unknown>,
  ) => string | number | undefined;
}

export function defineTool<T extends TObject>(
  config: ToolConfig<T>,
): AgentTool {
  const { modifiedSlide: getModifiedSlide, execute, ...rest } = config;

  const wrappedExecute = async (
    toolCallId: string,
    params: Static<T>,
    signal?: AbortSignal,
  ): Promise<ToolResult> => {
    if (signal?.aborted) {
      throw new DOMException("Tool execution aborted", "AbortError");
    }
    const result = await execute(toolCallId, params, signal);
    if (!getModifiedSlide) return result;
    const first = result.content[0];
    if (!first || first.type !== "text") return result;

    try {
      const parsed = JSON.parse(first.text);
      if (parsed.error) return result;

      const modifiedSlide = getModifiedSlide(params, parsed);
      if (typeof modifiedSlide === "string") {
        parsed._modifiedSlideId = modifiedSlide;
        return toolSuccess(parsed);
      }
      if (typeof modifiedSlide === "number") {
        parsed._modifiedSlide = modifiedSlide;
        return toolSuccess(parsed);
      }
    } catch {
      // Invalid JSON, return as-is
    }
    return result;
  };

  return { ...rest, execute: wrappedExecute } as unknown as AgentTool;
}

const PRESERVED_RESULT_METADATA_KEYS = new Set([
  "_modifiedSlide",
  "_modifiedSlideId",
  "slideId",
  "slideIndex",
  "positionOneIndexed",
  "originalSlideId",
  "replacementSlideId",
  "sourceSlideId",
  "sourceSlideIndex",
  "newSlideId",
  "newSlideIds",
  "newSlideIndex",
  "directoryVersion",
  "directoryChanged",
  "inputDirectoryChanged",
  "relocated",
  "usedLegacyIndex",
  "mutationCompleted",
  "mutationState",
  "code",
]);

const PRESERVED_ERROR_METADATA_KEYS = new Set([
  ...PRESERVED_RESULT_METADATA_KEYS,
  "expectedVersion",
  "currentVersion",
]);

function preservedMetadata(
  value: unknown,
  keys: ReadonlySet<string>,
): Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return {};
  }
  const record = value as Record<string, unknown>;
  const metadata: Record<string, unknown> = {};
  for (const key of keys) {
    const field = record[key];
    if (field !== undefined) metadata[key] = field;
  }
  return metadata;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error && error.message.length > 0) return error.message;
  if (typeof error === "string") return error;
  if (typeof error === "object" && error !== null) {
    const message = (error as { message?: unknown }).message;
    if (typeof message === "string" && message.length > 0) return message;
  }
  return "Unknown tool error";
}

export function toolSuccess(
  data: unknown,
  maxBytes = TOOL_RESULT_MAX_BYTES,
): ToolResult {
  const result =
    typeof data === "object" && data !== null ? { ...data } : { result: data };
  const serialized = JSON.stringify(result);
  const capped = capText(serialized, maxBytes);
  if (capped.truncated) {
    const limit = Math.max(512, Math.floor(maxBytes));
    const metadata = preservedMetadata(result, PRESERVED_RESULT_METADATA_KEYS);
    let previewLimit = Math.max(256, limit - 512);
    let preview = capText(serialized, previewLimit);
    let output = JSON.stringify({
      success: true,
      truncated: true,
      ...metadata,
      resultPreview: preview.text,
      omittedBytes: preview.omittedBytes,
      nextAction:
        "Narrow the request with slide_id/slide_index/shape_id/offset/limit or request a smaller field set.",
    });
    while (
      new TextEncoder().encode(output).byteLength > limit &&
      previewLimit > 256
    ) {
      previewLimit = Math.max(256, Math.floor(previewLimit * 0.75));
      preview = capText(serialized, previewLimit);
      output = JSON.stringify({
        success: true,
        truncated: true,
        ...metadata,
        resultPreview: preview.text,
        omittedBytes: preview.omittedBytes,
        nextAction:
          "Narrow the request with slide_id/slide_index/shape_id/offset/limit or request a smaller field set.",
      });
    }
    return {
      content: [
        {
          type: "text",
          text: output,
        },
      ],
      details: undefined,
    };
  }
  return {
    content: [{ type: "text", text: serialized }],
    details: undefined,
  };
}

export function toolError(
  error: unknown,
  maxBytesOrMetadata: number | object = TOOL_RESULT_MAX_BYTES,
): ToolResult {
  const maxBytes =
    typeof maxBytesOrMetadata === "number"
      ? maxBytesOrMetadata
      : TOOL_RESULT_MAX_BYTES;
  const message = errorMessage(error);
  const metadata = {
    ...preservedMetadata(error, PRESERVED_ERROR_METADATA_KEYS),
    ...(typeof maxBytesOrMetadata === "number"
      ? {}
      : preservedMetadata(maxBytesOrMetadata, PRESERVED_ERROR_METADATA_KEYS)),
  };
  const serialized = JSON.stringify({
    success: false,
    error: message,
    ...metadata,
  });
  const capped = capText(serialized, maxBytes);
  if (!capped.truncated) {
    return {
      content: [{ type: "text", text: serialized }],
      details: undefined,
    };
  }

  const limit = Math.max(512, Math.floor(maxBytes));
  let previewLimit = Math.max(256, limit - 512);
  let preview = capText(message, previewLimit);
  let output = JSON.stringify({
    success: false,
    truncated: true,
    error: "Tool error output truncated.",
    ...metadata,
    errorPreview: preview.text,
    omittedBytes: preview.omittedBytes,
  });
  while (
    new TextEncoder().encode(output).byteLength > limit &&
    previewLimit > 256
  ) {
    previewLimit = Math.max(256, Math.floor(previewLimit * 0.75));
    preview = capText(message, previewLimit);
    output = JSON.stringify({
      success: false,
      truncated: true,
      error: "Tool error output truncated.",
      ...metadata,
      errorPreview: preview.text,
      omittedBytes: preview.omittedBytes,
    });
  }
  return {
    content: [{ type: "text", text: output }],
    details: undefined,
  };
}

export function toolText(text: string): ToolResult {
  const marker = `\n\n[Output truncated. Use a narrower scope or pagination.]`;
  const markerBytes = new TextEncoder().encode(marker).byteLength;
  const capped = capText(text, TOOL_RESULT_MAX_BYTES - markerBytes);
  const output = capped.truncated ? `${capped.text}${marker}` : capped.text;
  return {
    content: [{ type: "text", text: output }],
    details: undefined,
  };
}

export async function toolImage(
  base64Data: string,
  mimeType: string,
  metadata?: unknown,
): Promise<ToolResult> {
  const resized = await resizeImage(base64Data, mimeType);
  const content: ToolResult["content"] = [
    {
      type: "image" as const,
      data: resized.data,
      mimeType: resized.mimeType,
    },
  ];
  if (metadata !== undefined) {
    content.push({
      type: "text" as const,
      text: JSON.stringify(metadata),
    });
  }
  return {
    content,
    details: undefined,
  };
}
