import { describe, expect, it, vi } from "vitest";
import {
  cleanupSlideMasters,
  type MasterCleanupMutationError,
} from "../src/lib/pptx/master-cleanup";
import { getSlideDirectoryVersion } from "../src/lib/pptx/slide-directory";

function createContext(injectConcurrentBeforeDelete = false) {
  const primaryTitle = { id: "layout-primary", name: "Title" };
  const primaryBody = { id: "layout-body", name: "Body" };
  const orphanLayout = { id: "layout-orphan", name: "Body" };
  const makeSlide = (id: string, layout = primaryTitle) => ({
    id,
    layout: { ...layout, load: vi.fn() },
    applyLayout: vi.fn(),
  });
  const items = [makeSlide("slide-1"), makeSlide("slide-2", orphanLayout)];
  let loadCount = 0;
  const deletedIds: string[] = [];
  const slides = {
    items,
    load: vi.fn(() => {
      loadCount++;
      if (
        injectConcurrentBeforeDelete &&
        loadCount === 5 &&
        !items.some((slide) => slide.id === "concurrent-slide")
      ) {
        items.push(makeSlide("concurrent-slide"));
      }
    }),
    add: vi.fn(() => {
      items.push(makeSlide("temporary-slide", orphanLayout));
    }),
    getItem: vi.fn((slideId: string) => ({
      delete: vi.fn(() => {
        deletedIds.push(slideId);
        const index = items.findIndex((slide) => slide.id === slideId);
        if (index >= 0) items.splice(index, 1);
      }),
    })),
  };
  const primaryMaster = {
    layouts: {
      items: [primaryTitle, primaryBody],
      load: vi.fn(),
    },
  };
  const orphanMaster = {
    layouts: { items: [orphanLayout], load: vi.fn() },
  };
  const context = {
    presentation: {
      slides,
      slideMasters: {
        items: [primaryMaster, orphanMaster],
        load: vi.fn(),
      },
    },
    sync: vi.fn(async () => undefined),
  };
  return { context, slides, items, deletedIds };
}

describe("cleanupSlideMasters stable temporary slide handling", () => {
  it("captures the temporary slide ID and deletes that exact slide", async () => {
    const { context, slides, items, deletedIds } = createContext();
    const initialVersion = getSlideDirectoryVersion(["slide-1", "slide-2"]);

    const receipt = await cleanupSlideMasters(
      context as unknown as PowerPoint.RequestContext,
      initialVersion,
    );

    expect(slides.getItem).toHaveBeenCalledWith("temporary-slide");
    expect(deletedIds).toEqual(["temporary-slide"]);
    expect(items.map((slide) => slide.id)).toEqual(["slide-1", "slide-2"]);
    expect(receipt).toMatchObject({
      directoryVersion: initialVersion,
      orphanedMasterCount: 1,
      temporarySlidesCreated: 1,
    });
  });

  it("fails closed when the directory changes before deletion", async () => {
    const { context, slides, items, deletedIds } = createContext(true);
    const initialVersion = getSlideDirectoryVersion(["slide-1", "slide-2"]);

    await expect(
      cleanupSlideMasters(
        context as unknown as PowerPoint.RequestContext,
        initialVersion,
      ),
    ).rejects.toMatchObject({
      code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
      cleanupPhase: "before_delete_temporary_slide",
      temporarySlideId: "temporary-slide",
      mutationCompleted: true,
    } satisfies Partial<MasterCleanupMutationError>);
    expect(slides.getItem).not.toHaveBeenCalled();
    expect(deletedIds).toEqual([]);
    expect(items.map((slide) => slide.id)).toEqual([
      "slide-1",
      "slide-2",
      "temporary-slide",
      "concurrent-slide",
    ]);
  });
});
