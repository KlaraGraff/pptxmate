import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import {
  applyPlainTextChange,
  hashPlainText,
  inspectPlainTextSelection,
  MAX_BATCH_PLAIN_TEXT_WRITE_BYTES,
  MAX_PLAIN_TEXT_WRITE_BYTES,
  utf8ByteLength,
} from "../pptx/text-xml";
import { findShapeById } from "../pptx/xml-utils";
import { unpackSlideZipResult } from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

export const updateSlideTextTool = defineTool({
  name: "update_slide_text",
  label: "Update Slide Text",
  description:
    "Apply multiple plain-text replacements or appended translations on one " +
    "slide in a single OOXML import. Text-body, paragraph, and first-run styling " +
    `is reused. Each text is limited to ${MAX_PLAIN_TEXT_WRITE_BYTES} UTF-8 ` +
    `bytes; aggregate text plus exact guards is limited to ${MAX_BATCH_PLAIN_TEXT_WRITE_BYTES}. ` +
    "Use expected_text_hash to keep guarded batches compact. Mixed runs may be " +
    "simplified; use edit_slide_text with code for detailed formatting changes.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    updates: Type.Array(
      Type.Object({
        shape_id: Type.String({
          description: "Stable shape ID from list_slide_shapes/read_slide_text",
        }),
        text: Type.String({
          description: `Replacement or appended text (maximum ${MAX_PLAIN_TEXT_WRITE_BYTES} UTF-8 bytes per shape)`,
          maxLength: MAX_PLAIN_TEXT_WRITE_BYTES,
        }),
        mode: Type.Optional(
          Type.Union([Type.Literal("replace"), Type.Literal("append")], {
            description: "Default replace",
          }),
        ),
        expected_text_hash: Type.Optional(
          Type.String({
            description:
              "Optional shapeTextHash from read_slide_text; aborts the entire batch if stale",
            maxLength: 32,
          }),
        ),
        expected_text: Type.Optional(
          Type.String({
            description:
              "Optional exact original shape text; aborts the entire batch if stale",
            maxLength: MAX_PLAIN_TEXT_WRITE_BYTES,
          }),
        ),
      }),
      { minItems: 1, maxItems: 50 },
    ),
    explanation: Type.Optional(
      Type.String({
        description: "Brief description (max 80 chars)",
        maxLength: 80,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      let batchBytes = 0;
      for (const update of params.updates) {
        const updateBytes = utf8ByteLength(update.text);
        if (updateBytes > MAX_PLAIN_TEXT_WRITE_BYTES) {
          throw new Error(
            `Text for shape "${update.shape_id}" is ${updateBytes} UTF-8 bytes; the per-shape limit is ${MAX_PLAIN_TEXT_WRITE_BYTES}. Use read_slide_text pagination and edit_slide_text range writes.`,
          );
        }
        if (
          update.expected_text !== undefined &&
          utf8ByteLength(update.expected_text) > MAX_PLAIN_TEXT_WRITE_BYTES
        ) {
          throw new Error(
            `expected_text for shape "${update.shape_id}" exceeds ${MAX_PLAIN_TEXT_WRITE_BYTES} UTF-8 bytes; use expected_text_hash instead.`,
          );
        }
        batchBytes += updateBytes;
        if (update.expected_text !== undefined) {
          batchBytes += utf8ByteLength(update.expected_text);
        }
      }
      if (batchBytes > MAX_BATCH_PLAIN_TEXT_WRITE_BYTES) {
        throw new Error(
          `Batch text and guard payload is ${batchBytes} UTF-8 bytes; the per-call limit is ${MAX_BATCH_PLAIN_TEXT_WRITE_BYTES}. Split the updates or use expected_text_hash instead of expected_text.`,
        );
      }

      const target = toSlideTargetReference(params);
      const zipValue = await safeRun(async (context) => {
        return withSlideZip(context, target, async ({ zip, markDirty }) => {
          const slideFile = zip.file("ppt/slides/slide1.xml");
          if (!slideFile) throw new Error("Slide XML not found in archive");
          const xml = await slideFile.async("string");
          const doc = new DOMParser().parseFromString(xml, "text/xml");

          const resolvedUpdates = params.updates.map((update) => {
            const shape = findShapeById(doc, update.shape_id);
            if (!shape) {
              throw new Error(
                `Shape with id "${update.shape_id}" not found on the targeted slide`,
              );
            }
            const selection = inspectPlainTextSelection(shape);
            if (
              update.expected_text !== undefined &&
              update.expected_text !== selection.text
            ) {
              throw new Error(
                `Text guard mismatch for shape "${update.shape_id}" (expected ${hashPlainText(update.expected_text)}, actual ${selection.textHash}). Re-read the shape before writing.`,
              );
            }
            if (
              update.expected_text_hash !== undefined &&
              update.expected_text_hash !== selection.textHash
            ) {
              throw new Error(
                `Text guard mismatch for shape "${update.shape_id}" (expected ${update.expected_text_hash}, actual ${selection.textHash}). Re-read the shape before writing.`,
              );
            }
            return { shape, update };
          });

          for (const { shape, update } of resolvedUpdates) {
            applyPlainTextChange(
              doc,
              shape,
              update.text,
              update.mode ?? "replace",
            );
          }

          zip.file(
            "ppt/slides/slide1.xml",
            new XMLSerializer().serializeToString(doc),
          );
          markDirty();
        });
      });
      const { metadata } = unpackSlideZipResult(zipValue, target);
      return toolSuccess({
        success: true,
        ...metadata,
        updatedShapeIds: params.updates.map((update) => update.shape_id),
      });
    } catch (error) {
      const normalized = toSlideMutationNotStartedError(
        error,
        "Failed to update slide text",
      );
      return toolError(normalized.message, normalized);
    }
  },
  modifiedSlide: (params, result) =>
    result && typeof result.slideId === "string"
      ? result.slideId
      : (params.slide_id ?? params.slide_index),
});
