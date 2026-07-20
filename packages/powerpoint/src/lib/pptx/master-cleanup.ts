import {
  createSlideDirectorySnapshot,
  loadSlideDirectory,
  type SlideDirectorySnapshot,
} from "./slide-directory";

/* global PowerPoint */

export interface MasterCleanupReceipt {
  directoryVersion: string;
  reassignedSlideCount: number;
  orphanedMasterCount: number;
  temporarySlidesCreated: number;
}

export class MasterCleanupMutationError extends Error {
  readonly mutationCompleted = true;

  constructor(
    readonly code:
      | "SLIDE_DIRECTORY_CHANGED_DURING_WRITE"
      | "MASTER_CLEANUP_TEMP_SLIDE_FAILED",
    readonly cleanupPhase: string,
    readonly expectedVersion?: string,
    readonly currentVersion?: string,
    readonly temporarySlideId?: string,
    detail?: string,
  ) {
    const versions =
      expectedVersion && currentVersion
        ? ` (expected ${expectedVersion}, current ${currentVersion})`
        : "";
    const temporary = temporarySlideId
      ? ` Temporary slide ID: ${temporarySlideId}.`
      : "";
    super(
      `[${code} mutationCompleted=true] Master cleanup stopped during ${cleanupPhase}${versions}.${temporary}${
        detail ? ` ${detail}` : ""
      } Do not replay the master edit; refresh list_slides and verify the presentation first.`,
    );
    this.name = "MasterCleanupMutationError";
  }
}

function assertDirectoryVersion(
  snapshot: SlideDirectorySnapshot,
  expectedVersion: string,
  phase: string,
  temporarySlideId?: string,
): void {
  if (snapshot.directoryVersion === expectedVersion) return;
  throw new MasterCleanupMutationError(
    "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
    phase,
    expectedVersion,
    snapshot.directoryVersion,
    temporarySlideId,
  );
}

function captureTemporarySlideId(
  before: SlideDirectorySnapshot,
  after: SlideDirectorySnapshot,
): string {
  const beforeIds = new Set(before.slideIds);
  const afterIds = new Set(after.slideIds);
  const addedIds = after.slideIds.filter((slideId) => !beforeIds.has(slideId));
  const removedIds = before.slideIds.filter(
    (slideId) => !afterIds.has(slideId),
  );
  const existingOrderUnchanged = before.slideIds.every(
    (slideId, index) => after.slideIds[index] === slideId,
  );

  if (
    addedIds.length !== 1 ||
    removedIds.length !== 0 ||
    after.slideIds.length !== before.slideIds.length + 1 ||
    !existingOrderUnchanged ||
    after.slideIds[after.slideIds.length - 1] !== addedIds[0]
  ) {
    throw new MasterCleanupMutationError(
      "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
      "capture_temporary_slide",
      before.directoryVersion,
      after.directoryVersion,
      undefined,
      "The temporary slide could not be identified uniquely, so no slide was deleted.",
    );
  }
  return addedIds[0];
}

export async function cleanupSlideMasters(
  context: PowerPoint.RequestContext,
  expectedDirectoryVersion?: string,
): Promise<MasterCleanupReceipt> {
  const masters = context.presentation.slideMasters;
  const slides = context.presentation.slides;
  masters.load("items");
  slides.load("items/id");
  await context.sync();

  const initialDirectory = createSlideDirectorySnapshot(slides.items);
  if (expectedDirectoryVersion) {
    assertDirectoryVersion(
      initialDirectory,
      expectedDirectoryVersion,
      "start_cleanup",
    );
  }

  const emptyReceipt = (): MasterCleanupReceipt => ({
    directoryVersion: initialDirectory.directoryVersion,
    reassignedSlideCount: 0,
    orphanedMasterCount: Math.max(0, masters.items.length - 1),
    temporarySlidesCreated: 0,
  });
  if (masters.items.length <= 1 || slides.items.length === 0) {
    return emptyReceipt();
  }

  for (const master of masters.items) {
    master.layouts.load("items/name,items/id");
  }
  for (const slide of slides.items) {
    slide.layout.load("name,id");
  }
  await context.sync();

  const primaryLayoutId = slides.items[0].layout.id;
  const primaryMaster = masters.items.find((master) =>
    master.layouts.items.some((layout) => layout.id === primaryLayoutId),
  );
  if (!primaryMaster) return emptyReceipt();

  let reassignedSlideCount = 0;
  for (let index = 1; index < slides.items.length; index++) {
    const layoutName = slides.items[index].layout.name;
    const matchingLayout = primaryMaster.layouts.items.find(
      (layout) => layout.name === layoutName,
    );
    if (matchingLayout) {
      slides.items[index].applyLayout(matchingLayout);
      reassignedSlideCount++;
    }
  }
  await context.sync();

  const postLayoutDirectory = await loadSlideDirectory(context);
  assertDirectoryVersion(
    postLayoutDirectory,
    initialDirectory.directoryVersion,
    "apply_primary_layouts",
  );

  const orphanedMasters = masters.items.filter(
    (master) => master !== primaryMaster,
  );
  let temporarySlidesCreated = 0;
  for (const master of orphanedMasters) {
    const firstLayout = master.layouts.items[0];
    if (!firstLayout) continue;

    const beforeAdd = await loadSlideDirectory(context);
    assertDirectoryVersion(
      beforeAdd,
      initialDirectory.directoryVersion,
      "before_temporary_slide",
    );

    try {
      slides.add({ layoutId: firstLayout.id });
      await context.sync();
    } catch (error) {
      throw new MasterCleanupMutationError(
        "MASTER_CLEANUP_TEMP_SLIDE_FAILED",
        "add_temporary_slide",
        beforeAdd.directoryVersion,
        undefined,
        undefined,
        error instanceof Error ? error.message : String(error),
      );
    }

    const afterAdd = await loadSlideDirectory(context);
    const temporarySlideId = captureTemporarySlideId(beforeAdd, afterAdd);
    temporarySlidesCreated++;

    const beforeDelete = await loadSlideDirectory(context);
    assertDirectoryVersion(
      beforeDelete,
      afterAdd.directoryVersion,
      "before_delete_temporary_slide",
      temporarySlideId,
    );

    try {
      slides.getItem(temporarySlideId).delete();
      await context.sync();
    } catch (error) {
      throw new MasterCleanupMutationError(
        "MASTER_CLEANUP_TEMP_SLIDE_FAILED",
        "delete_temporary_slide",
        afterAdd.directoryVersion,
        undefined,
        temporarySlideId,
        error instanceof Error ? error.message : String(error),
      );
    }

    const afterDelete = await loadSlideDirectory(context);
    assertDirectoryVersion(
      afterDelete,
      beforeAdd.directoryVersion,
      "after_delete_temporary_slide",
      temporarySlideId,
    );
  }

  const finalDirectory = await loadSlideDirectory(context);
  assertDirectoryVersion(
    finalDirectory,
    initialDirectory.directoryVersion,
    "finish_cleanup",
  );
  return {
    directoryVersion: finalDirectory.directoryVersion,
    reassignedSlideCount,
    orphanedMasterCount: orphanedMasters.length,
    temporarySlidesCreated,
  };
}
