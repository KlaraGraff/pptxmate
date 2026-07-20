import { Type } from "@sinclair/typebox";
import {
  createSlideDirectorySnapshot,
  resolveSlideTarget,
  slideTargetParameterProperties,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

export const listSlideShapesTool = defineTool({
  name: "list_slide_shapes",
  label: "List Slide Shapes",
  description:
    "List shapes on a slide with stable IDs. By default this is a compact text " +
    "index (id/name/type only); request include_geometry=true for positions. " +
    "Call this to discover shape IDs before using read_slide_text or edit_slide_text — " +
    "always use the shape ID (stable, locale-independent), never guess shape names.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    include_geometry: Type.Optional(
      Type.Boolean({
        description:
          "Include left/top/width/height. Default false; use only for layout work.",
      }),
    ),
    offset: Type.Optional(
      Type.Number({ description: "0-based shape offset. Default 0." }),
    ),
    limit: Type.Optional(
      Type.Number({ description: "Maximum shapes to return. Default 100." }),
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
        const resolved = resolveSlideTarget(
          directory,
          toSlideTargetReference(params),
        );
        const slide = slides.getItem(resolved.slideId);
        const shapes = slide.shapes;
        shapes.load(
          params.include_geometry
            ? "items/id,items/name,items/type,items/left,items/top,items/width,items/height"
            : "items/id,items/name,items/type",
        );
        await context.sync();

        const offset = Math.max(0, Math.floor(params.offset ?? 0));
        const limit = Math.min(
          200,
          Math.max(1, Math.floor(params.limit ?? 100)),
        );
        const all = shapes.items.map((s) => {
          const base = { id: s.id, name: s.name, type: s.type };
          if (!params.include_geometry) return base;
          return {
            ...base,
            left: s.left,
            top: s.top,
            width: s.width,
            height: s.height,
          };
        });
        const items = all.slice(offset, offset + limit);
        return {
          schemaVersion: 3,
          slideId: resolved.slideId,
          slideIndex: resolved.slideIndex,
          positionOneIndexed: resolved.slideIndex + 1,
          directoryVersion: directory.directoryVersion,
          directoryChanged: resolved.directoryChanged,
          relocated: resolved.indexMismatch,
          usedLegacyIndex: resolved.usedLegacyIndex,
          items,
          page: {
            offset,
            limit,
            total: all.length,
            nextOffset:
              offset + items.length < all.length ? offset + items.length : null,
          },
          omittedFields: params.include_geometry
            ? []
            : ["left", "top", "width", "height"],
        };
      });

      return toolSuccess({
        success: true,
        result: result.items,
        schemaVersion: result.schemaVersion,
        slideId: result.slideId,
        slideIndex: result.slideIndex,
        positionOneIndexed: result.positionOneIndexed,
        directoryVersion: result.directoryVersion,
        directoryChanged: result.directoryChanged,
        relocated: result.relocated,
        usedLegacyIndex: result.usedLegacyIndex,
        page: result.page,
        omittedFields: result.omittedFields,
      });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Failed to list slide shapes";
      return toolError(message);
    }
  },
});
