import { Type } from "@sinclair/typebox";
import {
  assertSlideDirectoryVersion,
  createSlideDirectorySnapshot,
  resolveSlideIds,
} from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

interface ShapeInfo {
  id: string;
  name: string;
  left: number;
  top: number;
  w: number;
  h: number;
}

interface VerifyResult {
  shapes: ShapeInfo[];
  overflows: Array<{
    shape: ShapeInfo;
    overflowBy: number;
    [key: string]: unknown;
  }>;
  overlaps: Array<{
    shapeA: ShapeInfo;
    shapeB: ShapeInfo;
    overlapX: number;
    overlapY: number;
  }>;
}

function verifyShapes(
  shapes: {
    id: string;
    name: string;
    left: number;
    top: number;
    width: number;
    height: number;
  }[],
  slideWidth: number,
  slideHeight: number,
): VerifyResult {
  const infos: ShapeInfo[] = [];
  const overflows: VerifyResult["overflows"] = [];
  const overlaps: VerifyResult["overlaps"] = [];

  for (const shape of shapes) {
    const info: ShapeInfo = {
      id: shape.id,
      name: shape.name,
      left: shape.left,
      top: shape.top,
      w: shape.width,
      h: shape.height,
    };
    infos.push(info);

    if (shape.top + shape.height > slideHeight) {
      overflows.push({
        shape: info,
        bottom: shape.top + shape.height,
        slideH: slideHeight,
        overflowBy: shape.top + shape.height - slideHeight,
      });
    }
    if (shape.left + shape.width > slideWidth) {
      overflows.push({
        shape: info,
        right: shape.left + shape.width,
        slideW: slideWidth,
        overflowBy: shape.left + shape.width - slideWidth,
      });
    }
  }

  for (let i = 0; i < infos.length; i++) {
    for (let j = i + 1; j < infos.length; j++) {
      const a = infos[i];
      const b = infos[j];

      if (
        a.left < b.left + b.w &&
        a.left + a.w > b.left &&
        a.top < b.top + b.h &&
        a.top + a.h > b.top
      ) {
        const overlapX =
          Math.min(a.left + a.w, b.left + b.w) - Math.max(a.left, b.left);
        const overlapY =
          Math.min(a.top + a.h, b.top + b.h) - Math.max(a.top, b.top);
        overlaps.push({ shapeA: a, shapeB: b, overlapX, overlapY });
      }
    }
  }

  return { shapes: infos, overflows, overlaps };
}

export const verifySlidesTool = defineTool({
  name: "verify_slides",
  label: "Verify Slides",
  description:
    "Verify slides for overlapping shapes and out-of-bounds positioning. " +
    "Returns a compact issue list by default; set include_shapes=true for the " +
    "legacy full geometry report.",
  parameters: Type.Object({
    only_issues: Type.Optional(
      Type.Boolean({
        description:
          "Return only overflow/overlap issues (default true). Set false for the legacy full shape report.",
      }),
    ),
    include_shapes: Type.Optional(
      Type.Boolean({
        description:
          "Include every shape's geometry. Default false; use only when needed.",
      }),
    ),
    slide_indices: Type.Optional(
      Type.Array(Type.Number(), {
        description:
          "Optional 0-based slide indices to verify; default all slides.",
        maxItems: 200,
      }),
    ),
    slide_ids: Type.Optional(
      Type.Array(Type.String(), {
        description:
          "Optional stable slide IDs from list_slides. Preferred over slide_indices.",
        maxItems: 200,
      }),
    ),
    directory_version: Type.Optional(
      Type.String({
        description:
          "Directory version from list_slides. Stable IDs remain authoritative; index-only requests fail if this version is stale.",
        minLength: 1,
        maxLength: 80,
      }),
    ),
    max_issues: Type.Optional(
      Type.Number({ description: "Maximum issues returned. Default 100." }),
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

        const pageSetup = context.presentation.pageSetup;
        pageSetup.load(["slideWidth", "slideHeight"]);
        await context.sync();

        const slideWidth = pageSetup.slideWidth;
        const slideHeight = pageSetup.slideHeight;
        const directory = createSlideDirectorySnapshot(slides.items);
        if (!params.slide_ids || params.slide_ids.length === 0) {
          assertSlideDirectoryVersion(directory, params.directory_version);
        }
        const resolvedTargets = resolveSlideIds(
          directory,
          params.slide_ids,
          params.slide_indices,
        );
        const results: Array<
          VerifyResult & {
            slideId: string;
            slideIndex: number;
            positionOneIndexed: number;
          }
        > = [];
        const maxIssues = Math.min(
          500,
          Math.max(1, Math.floor(params.max_issues ?? 100)),
        );
        let issueCount = 0;
        const targetSlideCount = resolvedTargets.ids.length;

        for (const slideId of resolvedTargets.ids) {
          const slideIndex = directory.indexById.get(slideId) as number;
          const slide = slides.getItem(slideId);
          const shapes = slide.shapes;
          shapes.load(
            "items/id,items/name,items/left,items/top,items/width,items/height",
          );
          await context.sync();

          const verified = verifyShapes(shapes.items, slideWidth, slideHeight);
          results.push({
            ...verified,
            slideId,
            slideIndex,
            positionOneIndexed: slideIndex + 1,
          });
          issueCount += verified.overflows.length + verified.overlaps.length;
        }

        const onlyIssues = params.only_issues !== false;
        if (!onlyIssues || params.include_shapes) {
          let remainingIssues = maxIssues;
          return {
            schemaVersion: 2,
            directoryVersion: directory.directoryVersion,
            directoryChanged:
              params.slide_ids !== undefined &&
              params.directory_version !== undefined &&
              params.directory_version !== directory.directoryVersion,
            usedLegacyIndices: resolvedTargets.usedLegacyIndices,
            relocated: resolvedTargets.relocatedIds.length > 0,
            relocatedSlideIds: resolvedTargets.relocatedIds,
            slides: results.map((slide) => {
              const overflows = slide.overflows.slice(0, remainingIssues);
              remainingIssues -= overflows.length;
              const overlaps = slide.overlaps.slice(0, remainingIssues);
              remainingIssues -= overlaps.length;
              return { ...slide, overflows, overlaps };
            }),
            checkedSlideCount: results.length,
            requestedSlideCount: targetSlideCount,
            issueCount,
            truncated: issueCount > maxIssues,
          };
        }

        const issues = results.flatMap((slide) => [
          ...slide.overflows.map((overflow) => ({
            slideId: slide.slideId,
            slideIndex: slide.slideIndex,
            positionOneIndexed: slide.positionOneIndexed,
            type: "overflow" as const,
            shapeId: overflow.shape.id,
            overflowBy: overflow.overflowBy,
          })),
          ...slide.overlaps.map((overlap) => ({
            slideId: slide.slideId,
            slideIndex: slide.slideIndex,
            positionOneIndexed: slide.positionOneIndexed,
            type: "overlap" as const,
            shapeIds: [overlap.shapeA.id, overlap.shapeB.id],
            overlapX: overlap.overlapX,
            overlapY: overlap.overlapY,
          })),
        ]);
        return {
          schemaVersion: 2,
          directoryVersion: directory.directoryVersion,
          directoryChanged:
            params.slide_ids !== undefined &&
            params.directory_version !== undefined &&
            params.directory_version !== directory.directoryVersion,
          usedLegacyIndices: resolvedTargets.usedLegacyIndices,
          relocated: resolvedTargets.relocatedIds.length > 0,
          relocatedSlideIds: resolvedTargets.relocatedIds,
          onlyIssues: true,
          checkedSlideCount: results.length,
          requestedSlideCount: targetSlideCount,
          issueCount,
          issues: issues.slice(0, maxIssues),
          truncated: issueCount > maxIssues,
        };
      });

      return toolSuccess({ success: true, result });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Failed to verify slides";
      return toolError(message);
    }
  },
});
