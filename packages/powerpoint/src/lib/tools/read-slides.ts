import { Type } from "@sinclair/typebox";
import { createSlideDirectorySnapshot } from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import {
  isSerializedJsonWithinBudget,
  TOOL_RESULT_MAX_BYTES,
} from "./result-budget";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

const MAX_SLIDES_PER_READ = 25;
const DEFAULT_PREVIEW_CHARS = 240;
const MAX_PREVIEW_CHARS = 500;
const TOTAL_PREVIEW_MAX_BYTES = 12 * 1024;

function buildPreview(
  texts: string[],
  maxChars: number,
  maxBytes: number,
): {
  textPreview: string;
  previewBytes: number;
  previewTruncated: boolean;
} {
  const normalized = texts.join("\n").replace(/\s+/g, " ").trim();
  const codePoints = Array.from(normalized);
  const encoder = new TextEncoder();
  const included: string[] = [];
  let previewBytes = 0;
  for (const codePoint of codePoints) {
    if (included.length >= maxChars) break;
    const bytes = encoder.encode(codePoint).byteLength;
    if (previewBytes + bytes > maxBytes) break;
    included.push(codePoint);
    previewBytes += bytes;
  }
  return {
    textPreview: included.join(""),
    previewBytes,
    previewTruncated: included.length < codePoints.length,
  };
}

export const readSlidesTool = defineTool({
  name: "read_slides",
  label: "Read Slide Previews",
  description:
    "Read compact plain-text previews for up to 25 slides by current slide ID. " +
    "Use list_slides first, then request only the slides needed. Each preview targets " +
    "240 characters by default within a shared UTF-8 budget and omits geometry, formatting, notes, screenshots, masters, and OOXML. " +
    "If hasMore is true, call read_slides again with remainingSlideIds. Use read_slide_texts to drill into one slide after locating it.",
  parameters: Type.Object({
    slide_ids: Type.Array(Type.String(), {
      description: "Slide IDs returned by list_slides. Maximum 25.",
      minItems: 1,
      maxItems: MAX_SLIDES_PER_READ,
    }),
    preview_chars: Type.Optional(
      Type.Number({
        description:
          "Maximum Unicode characters per slide preview. Default 240, max 500.",
      }),
    ),
    directory_version: Type.Optional(
      Type.String({
        description:
          "Directory version from list_slides. Slide IDs remain authoritative if the directory has since changed.",
        minLength: 1,
        maxLength: 80,
      }),
    ),
    explanation: Type.Optional(
      Type.String({
        description: "Brief description (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      const result = await safeRun(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();

        const directory = createSlideDirectorySnapshot(slides.items);
        const uniqueIds = Array.from(new Set(params.slide_ids));
        const requestedIds = uniqueIds.slice(0, MAX_SLIDES_PER_READ);
        const slideById = new Map(
          slides.items.map((slide, slideIndex) => [
            slide.id,
            { slide, slideIndex },
          ]),
        );
        const unknownIds = requestedIds.filter((id) => !slideById.has(id));
        if (unknownIds.length > 0) {
          throw new Error(
            `Unknown slide ID(s): ${unknownIds.join(", ")}. Call list_slides again to refresh the directory.`,
          );
        }

        const targets = requestedIds.map((slideId) => {
          const target = slideById.get(slideId)!;
          const shapes = target.slide.shapes;
          shapes.load("items/id");
          return { slideId, ...target, shapes };
        });
        await context.sync();

        const framesBySlide = targets.map((target) =>
          target.shapes.items.map((shape) => {
            const textFrame = shape.getTextFrameOrNullObject();
            textFrame.load("hasText");
            return textFrame;
          }),
        );
        await context.sync();

        const textRangesBySlide = framesBySlide.map((frames) =>
          frames.map((textFrame) => {
            if (textFrame.isNullObject || !textFrame.hasText) {
              return null;
            }
            const textRange = textFrame.textRange;
            textRange.load("text");
            return textRange;
          }),
        );
        await context.sync();

        const previewChars = Math.min(
          MAX_PREVIEW_CHARS,
          Math.max(
            40,
            Math.floor(params.preview_chars ?? DEFAULT_PREVIEW_CHARS),
          ),
        );
        const previewByteBudget = Math.max(
          256,
          Math.floor(TOTAL_PREVIEW_MAX_BYTES / Math.max(1, targets.length)),
        );
        const previewItems = targets.map((target, targetIndex) => {
          const frames = framesBySlide[targetIndex];
          const textRanges = textRangesBySlide[targetIndex];
          const texts = textRanges
            .filter((range): range is PowerPoint.TextRange => range !== null)
            .map((range) => range.text);
          const textFrameShapeCount = frames.filter(
            (textFrame) => !textFrame.isNullObject,
          ).length;

          return {
            slideId: target.slideId,
            slideIndex: target.slideIndex,
            positionOneIndexed: target.slideIndex + 1,
            ...buildPreview(texts, previewChars, previewByteBudget),
            shapeCount: target.shapes.items.length,
            textFrameShapeCount,
            nonTextFrameShapeCount:
              target.shapes.items.length - textFrameShapeCount,
          };
        });

        const buildResult = (
          items: typeof previewItems,
          remainingSlideIds: string[],
        ) => ({
          schemaVersion: 2,
          directoryVersion: directory.directoryVersion,
          directoryChanged:
            params.directory_version !== undefined &&
            params.directory_version !== directory.directoryVersion,
          items,
          requested: params.slide_ids.length,
          returned: items.length,
          duplicateSlideIdsOmitted: params.slide_ids.length - uniqueIds.length,
          slideIdsOmittedByLimit: uniqueIds.length - requestedIds.length,
          slideIdsOmittedByBudget: requestedIds.length - items.length,
          hasMore: remainingSlideIds.length > 0,
          remainingSlideIds,
          previewChars,
          previewByteBudget,
          previewScope: "plain-text-shape-frames",
          omittedFields: [
            "font",
            "fontSize",
            "color",
            "position",
            "geometry",
            "notes",
            "screenshots",
            "masters",
            "layouts",
            "rawOoxml",
            "tableCellText",
            "chartText",
            "groupChildText",
          ],
        });

        let result = buildResult([], uniqueIds);
        if (
          !isSerializedJsonWithinBudget(
            { success: true, result },
            TOOL_RESULT_MAX_BYTES,
          )
        ) {
          throw new Error(
            "The requested slide IDs exceed the read_slides output budget. Request fewer slide IDs.",
          );
        }

        for (let returned = 1; returned <= previewItems.length; returned++) {
          const candidate = buildResult(
            previewItems.slice(0, returned),
            uniqueIds.slice(returned),
          );
          if (
            !isSerializedJsonWithinBudget(
              { success: true, result: candidate },
              TOOL_RESULT_MAX_BYTES,
            )
          ) {
            break;
          }
          result = candidate;
        }

        return result;
      });

      return toolSuccess({ success: true, result });
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : "Failed to read slide previews";
      return toolError(message);
    }
  },
});
