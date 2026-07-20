import type JSZip from "jszip";
import { default as JSZipCtor } from "jszip";
import {
  assertSlideDirectoryVersion,
  createSlideDirectorySnapshot,
  loadSlideDirectory,
  resolveSlideTarget,
  SlideDirectoryChangedDuringWriteError,
  type SlideTargetReference,
  toSlideMutationNotStartedError,
  toSlideMutationUncertainError,
} from "./slide-directory";
import { extractExternalReferences, sanitizeXmlAmpersands } from "./xml-utils";

/* global PowerPoint */

export interface SlideZipArgs {
  zip: JSZip;
  markDirty: () => void;
}

export interface SlideZipResult<T> {
  result: T;
  originalSlideId: string;
  slideId: string;
  replacementSlideId: string | null;
  slideIndex: number;
  directoryVersion: string;
  directoryChanged: boolean;
  inputDirectoryChanged: boolean;
  relocated: boolean;
  usedLegacyIndex: boolean;
  mutationCompleted: boolean;
  mutationState: "not_started" | "completed";
}

let onlineRunQueue = Promise.resolve();

function getPlatform(): string | undefined {
  if (typeof Office === "undefined") return undefined;
  try {
    return Office.context?.platform as unknown as string;
  } catch {
    return undefined;
  }
}

function safeRun<T>(
  callback: (ctx: PowerPoint.RequestContext) => Promise<T>,
): Promise<T> {
  if (getPlatform() !== "OfficeOnline") {
    return PowerPoint.run(callback);
  }
  let hostRun: Promise<T>;
  const task: Promise<T> = onlineRunQueue
    .catch(() => {})
    .then(() => {
      hostRun = PowerPoint.run(callback);
      return Promise.race([
        hostRun,
        new Promise<T>((_, reject) =>
          setTimeout(
            () =>
              reject(
                toSlideMutationUncertainError(
                  new Error("Office.run timed out after 120s"),
                  "Office.run timed out and may still complete its slide write.",
                ),
              ),
            120_000,
          ),
        ),
      ]);
    });
  onlineRunQueue = task
    .then(() => hostRun)
    .catch(() => hostRun)
    .then(() => {})
    .catch(() => {});
  return task;
}

export { safeRun };

export function withSlideZip<T>(
  context: PowerPoint.RequestContext,
  slideIndex: number,
  callback: (args: SlideZipArgs) => Promise<T>,
): Promise<T>;
export function withSlideZip<T>(
  context: PowerPoint.RequestContext,
  target: SlideTargetReference,
  callback: (args: SlideZipArgs) => Promise<T>,
): Promise<SlideZipResult<T>>;
export async function withSlideZip<T>(
  context: PowerPoint.RequestContext,
  target: number | SlideTargetReference,
  callback: (args: SlideZipArgs) => Promise<T>,
): Promise<T | SlideZipResult<T>> {
  const legacyIndexCall = typeof target === "number";
  const targetReference: SlideTargetReference = legacyIndexCall
    ? { slide_index: target }
    : target;
  let hostWriteSyncStarted = false;

  try {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();

    const initialDirectory = createSlideDirectorySnapshot(slides.items);
    const resolvedTarget = resolveSlideTarget(
      initialDirectory,
      targetReference,
    );
    const originalSlideId = resolvedTarget.slideId;

    const exportResult = slides.getItem(originalSlideId).exportAsBase64();
    await context.sync();

    const zip = await JSZipCtor.loadAsync(exportResult.value, { base64: true });

    let dirty = false;
    const refsBefore = await extractExternalReferences(zip);

    const result = await callback({
      zip,
      markDirty: () => {
        dirty = true;
      },
    });

    if (dirty) {
      const slideFile = zip.file("ppt/slides/slide1.xml");
      if (slideFile) {
        const slideXml = await slideFile.async("string");
        const sanitized = sanitizeXmlAmpersands(slideXml);
        if (sanitized !== slideXml) {
          zip.file("ppt/slides/slide1.xml", sanitized);
        }
      }

      const refsAfter = await extractExternalReferences(zip);
      for (const ref of refsAfter) {
        if (!refsBefore.has(ref)) {
          throw new Error(`Adding external references is blocked (${ref})`);
        }
      }

      const modifiedBase64 = await zip.generateAsync({ type: "base64" });

      // Refresh selection at the commit boundary so a long callback does not
      // restore a selection that the user changed while it was running.
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load("items/id");
      const preWriteDirectory = await loadSlideDirectory(context);
      const selectedSlideIds = selectedSlides.items.map((slide) => slide.id);
      assertSlideDirectoryVersion(
        preWriteDirectory,
        initialDirectory.directoryVersion,
      );
      const preWriteTarget = resolveSlideTarget(preWriteDirectory, {
        slide_id: originalSlideId,
        directory_version: initialDirectory.directoryVersion,
      });
      const insertionTargetId =
        preWriteTarget.slideIndex > 0
          ? preWriteDirectory.slideIds[preWriteTarget.slideIndex - 1]
          : undefined;

      context.presentation.insertSlidesFromBase64(modifiedBase64, {
        targetSlideId: insertionTargetId,
      });
      slides.getItem(originalSlideId).delete();
      hostWriteSyncStarted = true;
      await context.sync();

      const postWriteDirectory = await loadSlideDirectory(context);
      const beforeIds = new Set(preWriteDirectory.slideIds);
      const afterIds = new Set(postWriteDirectory.slideIds);
      const addedIds = postWriteDirectory.slideIds.filter(
        (slideId) => !beforeIds.has(slideId),
      );
      const removedIds = preWriteDirectory.slideIds.filter(
        (slideId) => !afterIds.has(slideId),
      );
      if (
        addedIds.length !== 1 ||
        removedIds.length !== 1 ||
        removedIds[0] !== originalSlideId
      ) {
        throw new SlideDirectoryChangedDuringWriteError(
          preWriteDirectory.directoryVersion,
          postWriteDirectory.directoryVersion,
        );
      }

      const replacementSlideId = addedIds[0];
      const replacementSlideIndex =
        postWriteDirectory.indexById.get(replacementSlideId);
      if (replacementSlideIndex === undefined) {
        throw new Error("Replacement slide could not be located after import.");
      }

      const expectedPostWriteIds = [...preWriteDirectory.slideIds];
      expectedPostWriteIds[preWriteTarget.slideIndex] = replacementSlideId;
      if (
        postWriteDirectory.slideIds.length !== expectedPostWriteIds.length ||
        postWriteDirectory.slideIds.some(
          (slideId, index) => slideId !== expectedPostWriteIds[index],
        )
      ) {
        throw new SlideDirectoryChangedDuringWriteError(
          preWriteDirectory.directoryVersion,
          postWriteDirectory.directoryVersion,
        );
      }

      const restoredSelection = selectedSlideIds
        .map((slideId) =>
          slideId === originalSlideId ? replacementSlideId : slideId,
        )
        .filter((slideId) => postWriteDirectory.indexById.has(slideId));
      if (restoredSelection.length > 0) {
        context.presentation.setSelectedSlides(restoredSelection);
        await context.sync();

        // Selection restoration is another host round trip. Recheck only the
        // lightweight ID directory before reporting a deterministic success.
        const afterSelectionDirectory = await loadSlideDirectory(context);
        if (
          afterSelectionDirectory.directoryVersion !==
          postWriteDirectory.directoryVersion
        ) {
          throw new SlideDirectoryChangedDuringWriteError(
            postWriteDirectory.directoryVersion,
            afterSelectionDirectory.directoryVersion,
          );
        }
      }

      if (legacyIndexCall) return result;
      return {
        result,
        originalSlideId,
        slideId: replacementSlideId,
        replacementSlideId,
        slideIndex: replacementSlideIndex,
        directoryVersion: postWriteDirectory.directoryVersion,
        directoryChanged: true,
        inputDirectoryChanged: resolvedTarget.directoryChanged,
        relocated: resolvedTarget.indexMismatch,
        usedLegacyIndex: resolvedTarget.usedLegacyIndex,
        mutationCompleted: true,
        mutationState: "completed",
      };
    }

    if (legacyIndexCall) return result;
    return {
      result,
      originalSlideId,
      slideId: originalSlideId,
      replacementSlideId: null,
      slideIndex: resolvedTarget.slideIndex,
      directoryVersion: initialDirectory.directoryVersion,
      directoryChanged: resolvedTarget.directoryChanged,
      inputDirectoryChanged: resolvedTarget.directoryChanged,
      relocated: resolvedTarget.indexMismatch,
      usedLegacyIndex: resolvedTarget.usedLegacyIndex,
      mutationCompleted: false,
      mutationState: "not_started",
    };
  } catch (error) {
    if (hostWriteSyncStarted) {
      throw toSlideMutationUncertainError(
        error,
        "The slide replacement may have completed; verify the target before retrying.",
      );
    }
    throw toSlideMutationNotStartedError(
      error,
      "The slide replacement did not start.",
    );
  }
}
