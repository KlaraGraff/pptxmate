import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import {
  applyPlainTextRangeChange,
  inspectPlainTextSelection,
  MAX_PLAIN_TEXT_WRITE_BYTES,
  MAX_TEXT_RANGE_CHARACTERS,
  MAX_TEXT_RANGE_PARAGRAPHS,
  type PlainTextRange,
  utf8ByteLength,
} from "../pptx/text-xml";
import { findShapeById } from "../pptx/xml-utils";
import { unpackSlideZipResult } from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

function assertInputBytes(value: string): void {
  const bytes = utf8ByteLength(value);
  if (bytes > MAX_PLAIN_TEXT_WRITE_BYTES) {
    throw new Error(
      `Patch text is ${bytes} UTF-8 bytes; the per-call limit is ${MAX_PLAIN_TEXT_WRITE_BYTES}. Split the patch into smaller ranges.`,
    );
  }
}

function toRange(params: {
  paragraph_start: number;
  paragraph_end: number;
  char_start?: number;
  char_end?: number;
}): PlainTextRange {
  const hasCharStart = params.char_start !== undefined;
  const hasCharEnd = params.char_end !== undefined;
  if (hasCharStart !== hasCharEnd) {
    throw new Error("Character ranges require both char_start and char_end");
  }
  if (hasCharStart && hasCharEnd) {
    return {
      kind: "characters",
      paragraphStart: params.paragraph_start,
      paragraphEnd: params.paragraph_end,
      charStart: params.char_start as number,
      charEnd: params.char_end as number,
    };
  }
  return {
    kind: "paragraphs",
    paragraphStart: params.paragraph_start,
    paragraphEnd: params.paragraph_end,
  };
}

function verificationScope(
  range: PlainTextRange,
  text: string,
  mode: "replace" | "append",
): Record<string, number> {
  if (range.kind === "characters") {
    return {
      paragraph_start: range.paragraphStart,
      paragraph_end: range.paragraphEnd,
      char_start: range.charStart,
      char_end:
        mode === "append"
          ? range.charEnd + text.length
          : range.charStart + text.length,
    };
  }
  const lineCount = text.replace(/\r\n?/g, "\n").split("\n").length;
  return {
    paragraph_start: range.paragraphStart,
    paragraph_end:
      mode === "append"
        ? range.paragraphEnd + lineCount
        : range.paragraphStart + lineCount,
  };
}

function verificationReadArgs(
  scope: Record<string, number>,
): Record<string, number> {
  const paragraphStart = scope.paragraph_start;
  const paragraphEnd = scope.paragraph_end;
  const args: Record<string, number> = {
    paragraph_offset: paragraphStart,
    paragraph_limit: paragraphEnd - paragraphStart,
  };
  if (scope.char_start !== undefined) {
    args.char_offset = scope.char_start;
    args.paragraph_limit = 1;
  }
  return args;
}

export const patchSlideTextTool = defineTool({
  name: "patch_slide_text",
  label: "Patch Slide Text Range",
  description:
    "Apply a guarded text patch to a paragraph or character range only. " +
    "Use this for translation audit repairs after read_slide_text returned the " +
    "exact editScope and textHash. It cannot replace an entire shape. Re-read " +
    "the returned verificationScope before reporting the repair complete.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    shape_id: Type.String({
      description: "Stable shape ID from read_slide_text.",
    }),
    text: Type.String({
      description: `Replacement or appended text (maximum ${MAX_PLAIN_TEXT_WRITE_BYTES} UTF-8 bytes).`,
      maxLength: MAX_PLAIN_TEXT_WRITE_BYTES,
    }),
    mode: Type.Optional(
      Type.Union([Type.Literal("replace"), Type.Literal("append")], {
        description: "Range mode. Default replace.",
      }),
    ),
    paragraph_start: Type.Integer({
      description: "Inclusive paragraph start from read_slide_text editScope.",
      minimum: 0,
    }),
    paragraph_end: Type.Integer({
      description: `Exclusive paragraph end from editScope (maximum ${MAX_TEXT_RANGE_PARAGRAPHS} paragraphs).`,
      minimum: 1,
    }),
    char_start: Type.Optional(
      Type.Integer({
        description: "Inclusive character start from editScope.",
        minimum: 0,
      }),
    ),
    char_end: Type.Optional(
      Type.Integer({
        description: `Exclusive character end from editScope (maximum ${MAX_TEXT_RANGE_CHARACTERS} UTF-16 code units).`,
        minimum: 1,
      }),
    ),
    expected_text_hash: Type.String({
      description:
        "Required textHash from the exact read_slide_text editScope being patched.",
      maxLength: 32,
    }),
    explanation: Type.Optional(
      Type.String({
        description: "Brief description (max 80 chars)",
        maxLength: 80,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      assertInputBytes(params.text);
      const range = toRange(params);
      const mode = params.mode ?? "replace";
      const target = toSlideTargetReference(params);
      let patchResult:
        | {
            beforeTextHash: string;
            afterTextHash: string;
            nextCursor: {
              paragraph_offset: number;
              char_offset?: number;
            } | null;
          }
        | undefined;
      const zipValue = await safeRun(async (context) =>
        withSlideZip(context, target, async ({ zip, markDirty }) => {
          const slideFile = zip.file("ppt/slides/slide1.xml");
          if (!slideFile) throw new Error("Slide XML not found in archive");
          const xml = await slideFile.async("string");
          const doc = new DOMParser().parseFromString(xml, "text/xml");
          const shape = findShapeById(doc, params.shape_id);
          if (!shape) {
            throw new Error(
              `Shape with id "${params.shape_id}" was not found on the targeted slide. Re-read the slide before patching.`,
            );
          }
          const selection = inspectPlainTextSelection(shape, range);
          if (selection.textHash !== params.expected_text_hash) {
            throw new Error(
              `Text guard mismatch for the requested range (expected ${params.expected_text_hash}, actual ${selection.textHash}). Re-read the exact range before patching.`,
            );
          }
          patchResult = applyPlainTextRangeChange(
            doc,
            shape,
            params.text,
            range,
            mode,
          );
          zip.file(
            "ppt/slides/slide1.xml",
            new XMLSerializer().serializeToString(doc),
          );
          markDirty();
        }),
      );
      const { metadata } = unpackSlideZipResult(zipValue, target);
      if (!patchResult) throw new Error("Text patch did not produce a result");
      const resultVerificationScope = verificationScope(
        range,
        params.text,
        mode,
      );
      return toolSuccess({
        success: true,
        ...metadata,
        shapeId: params.shape_id,
        mode,
        scope:
          range.kind === "characters"
            ? {
                paragraph_start: range.paragraphStart,
                paragraph_end: range.paragraphEnd,
                char_start: range.charStart,
                char_end: range.charEnd,
              }
            : {
                paragraph_start: range.paragraphStart,
                paragraph_end: range.paragraphEnd,
              },
        beforeTextHash: patchResult.beforeTextHash,
        afterTextHash: patchResult.afterTextHash,
        verificationScope: resultVerificationScope,
        verificationReadArgs: verificationReadArgs(resultVerificationScope),
        nextCursor: patchResult.nextCursor,
        nextAction:
          "Re-read verificationScope with read_slide_text before reporting this translation repair complete.",
      });
    } catch (error) {
      const normalized = toSlideMutationNotStartedError(
        error,
        "Failed to patch slide text",
      );
      return toolError(normalized.message, normalized);
    }
  },
  modifiedSlide: (params, result) =>
    result && typeof result.slideId === "string"
      ? result.slideId
      : (params.slide_id ?? params.slide_index),
});
