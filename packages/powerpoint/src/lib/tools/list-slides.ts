import { Type } from "@sinclair/typebox";
import {
  assertSlideDirectoryVersion,
  createSlideDirectorySnapshot,
} from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

const MAX_SLIDES_PER_PAGE = 25;

export const listSlidesTool = defineTool({
  name: "list_slides",
  label: "List Slides",
  description:
    "List a lightweight, paginated slide directory in presentation order. " +
    "Returns slide ID, zero-based index, one-based position, and selection state only. " +
    "It deliberately does not read slide text, shapes, geometry, formatting, notes, screenshots, masters, or OOXML.",
  parameters: Type.Object({
    offset: Type.Optional(
      Type.Number({ description: "0-based slide offset. Default 0." }),
    ),
    limit: Type.Optional(
      Type.Number({
        description: "Maximum slides to return. Default 25, max 25.",
      }),
    ),
    directory_version: Type.Optional(
      Type.String({
        description:
          "Directory version from the previous list_slides page. If it changed, restart pagination from offset 0.",
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

        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load("items/id");
        await context.sync();

        const directory = createSlideDirectorySnapshot(slides.items);
        assertSlideDirectoryVersion(directory, params.directory_version);
        const offset = Math.max(0, Math.floor(params.offset ?? 0));
        const limit = Math.min(
          MAX_SLIDES_PER_PAGE,
          Math.max(1, Math.floor(params.limit ?? MAX_SLIDES_PER_PAGE)),
        );
        const selectedIds = new Set(
          selectedSlides.items.map((slide) => slide.id),
        );
        const items = slides.items
          .slice(offset, offset + limit)
          .map((slide, pageIndex) => {
            const slideIndex = offset + pageIndex;
            return {
              slideId: slide.id,
              slideIndex,
              positionOneIndexed: slideIndex + 1,
              selected: selectedIds.has(slide.id),
            };
          });
        const nextOffset =
          offset + items.length < slides.items.length
            ? offset + items.length
            : null;

        return {
          schemaVersion: 2,
          directoryVersion: directory.directoryVersion,
          items,
          page: {
            offset,
            limit,
            total: slides.items.length,
            returned: items.length,
            hasMore: nextOffset !== null,
            nextOffset,
          },
          ordering: "deck-order",
          omittedFields: [
            "text",
            "shapes",
            "geometry",
            "fonts",
            "colors",
            "notes",
            "screenshots",
            "masters",
            "layouts",
            "rawOoxml",
          ],
        };
      });

      return toolSuccess({ success: true, result });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Failed to list slides";
      return toolError(message);
    }
  },
});
