import { Type } from "@sinclair/typebox";
import {
  createSlideDirectorySnapshot,
  resolveSlideTarget,
  slideTargetParameterProperties,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import { defineTool, toolError, toolImage } from "./types";

/* global PowerPoint */

export const screenshotSlideTool = defineTool({
  name: "screenshot_slide",
  label: "Screenshot Slide",
  description:
    "Take a screenshot of a slide for visual verification of layout, positioning, and content.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    explanation: Type.Optional(
      Type.String({
        description: "Brief description of the action (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      const imageData = await safeRun(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();
        const directory = createSlideDirectorySnapshot(slides.items);
        const resolved = resolveSlideTarget(
          directory,
          toSlideTargetReference(params),
        );
        const imageResult = slides
          .getItem(resolved.slideId)
          .getImageAsBase64({ width: 960 });
        await context.sync();
        return {
          data: imageResult.value,
          slideId: resolved.slideId,
          slideIndex: resolved.slideIndex,
          positionOneIndexed: resolved.slideIndex + 1,
          directoryVersion: directory.directoryVersion,
          directoryChanged: resolved.directoryChanged,
          relocated: resolved.indexMismatch,
          usedLegacyIndex: resolved.usedLegacyIndex,
        };
      });

      return await toolImage(imageData.data, "image/png", {
        success: true,
        slideId: imageData.slideId,
        slideIndex: imageData.slideIndex,
        positionOneIndexed: imageData.positionOneIndexed,
        directoryVersion: imageData.directoryVersion,
        directoryChanged: imageData.directoryChanged,
        relocated: imageData.relocated,
        usedLegacyIndex: imageData.usedLegacyIndex,
      });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Failed to screenshot slide";
      return toolError(message);
    }
  },
});
