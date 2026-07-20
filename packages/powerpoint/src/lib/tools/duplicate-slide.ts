import { Type } from "@sinclair/typebox";
import {
  assertSlideDirectoryVersion,
  createSlideDirectorySnapshot,
  loadSlideDirectory,
  resolveSlideTarget,
  SlideDirectoryChangedDuringWriteError,
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideMutationUncertainError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

export const duplicateSlideTool = defineTool({
  name: "duplicate_slide",
  label: "Duplicate Slide",
  description:
    "Duplicate a slide by stable slide ID. The copy is inserted immediately after the original and the refreshed ID directory is returned.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    explanation: Type.Optional(
      Type.String({
        description: "Brief description (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    let hostWriteSyncStarted = false;
    try {
      const result = await safeRun(async (context) => {
        const slides = context.presentation.slides;
        slides.load("items/id");
        await context.sync();

        const initialDirectory = createSlideDirectorySnapshot(slides.items);
        const resolved = resolveSlideTarget(
          initialDirectory,
          toSlideTargetReference(params),
        );

        const exported = slides.getItem(resolved.slideId).exportAsBase64();
        await context.sync();

        const preWriteDirectory = await loadSlideDirectory(context);
        assertSlideDirectoryVersion(
          preWriteDirectory,
          initialDirectory.directoryVersion,
        );
        resolveSlideTarget(preWriteDirectory, { slide_id: resolved.slideId });

        context.presentation.insertSlidesFromBase64(exported.value, {
          targetSlideId: resolved.slideId,
        });
        hostWriteSyncStarted = true;
        await context.sync();

        const postWriteDirectory = await loadSlideDirectory(context);
        const beforeSet = new Set(preWriteDirectory.slideIds);
        const addedIds = postWriteDirectory.slideIds.filter(
          (id) => !beforeSet.has(id),
        );
        if (addedIds.length !== 1) {
          throw new SlideDirectoryChangedDuringWriteError(
            preWriteDirectory.directoryVersion,
            postWriteDirectory.directoryVersion,
          );
        }
        const newSlideId = addedIds[0];
        const sourceSlideIndex = preWriteDirectory.indexById.get(
          resolved.slideId,
        );
        if (sourceSlideIndex === undefined) {
          throw new SlideDirectoryChangedDuringWriteError(
            preWriteDirectory.directoryVersion,
            postWriteDirectory.directoryVersion,
          );
        }
        const expectedIds = [...preWriteDirectory.slideIds];
        expectedIds.splice(sourceSlideIndex + 1, 0, newSlideId);
        if (
          postWriteDirectory.slideIds.length !== expectedIds.length ||
          postWriteDirectory.slideIds.some(
            (slideId, index) => slideId !== expectedIds[index],
          )
        ) {
          throw new SlideDirectoryChangedDuringWriteError(
            preWriteDirectory.directoryVersion,
            postWriteDirectory.directoryVersion,
          );
        }
        const newSlideIndex = postWriteDirectory.indexById.get(newSlideId);
        if (newSlideIndex === undefined) {
          throw new SlideDirectoryChangedDuringWriteError(
            preWriteDirectory.directoryVersion,
            postWriteDirectory.directoryVersion,
          );
        }
        return {
          sourceSlideId: resolved.slideId,
          sourceSlideIndex,
          newSlideId,
          newSlideIndex,
          positionOneIndexed: newSlideIndex + 1,
          directoryVersion: postWriteDirectory.directoryVersion,
          directoryChanged: true,
          inputDirectoryChanged: resolved.directoryChanged,
          relocated: resolved.indexMismatch,
          usedLegacyIndex: resolved.usedLegacyIndex,
          mutationCompleted: true,
          mutationState: "completed" as const,
        };
      });

      return toolSuccess({ success: true, ...result });
    } catch (error) {
      const normalized = hostWriteSyncStarted
        ? toSlideMutationUncertainError(
            error,
            "The slide duplication may have completed; refresh list_slides before retrying.",
          )
        : toSlideMutationNotStartedError(error, "Failed to duplicate slide");
      return toolError(normalized.message, normalized);
    }
  },
  modifiedSlide: (_params, result) =>
    result && typeof result.newSlideId === "string"
      ? result.newSlideId
      : undefined,
});
