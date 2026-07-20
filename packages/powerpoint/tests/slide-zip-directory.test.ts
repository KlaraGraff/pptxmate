import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  getSlideDirectoryVersion,
  SlideDirectoryChangedError,
} from "../src/lib/pptx/slide-directory";
import { withSlideZip } from "../src/lib/pptx/slide-zip";

interface FakeState {
  ids: string[];
  selectedIds: string[];
  exportedIds: string[];
  deletedIds: string[];
  insertionTargets: Array<string | undefined>;
  nextReplacement: number;
  syncCount: number;
  onSync?: (syncCount: number) => void;
}

function createFakeContext(initialIds: string[], selectedIds: string[] = []) {
  const state: FakeState = {
    ids: [...initialIds],
    selectedIds: [...selectedIds],
    exportedIds: [],
    deletedIds: [],
    insertionTargets: [],
    nextReplacement: 1,
    syncCount: 0,
  };

  const getSlide = (id: string) => ({
    id,
    exportAsBase64: () => {
      state.exportedIds.push(id);
      const zip = new JSZip();
      zip.file("ppt/slides/slide1.xml", `<p:sld data-slide="${id}"/>`);
      return { value: zip.generateAsync({ type: "base64" }) };
    },
    delete: () => {
      const index = state.ids.indexOf(id);
      if (index >= 0) state.ids.splice(index, 1);
      state.deletedIds.push(id);
    },
  });

  const slides = {
    load: () => undefined,
    get items() {
      return state.ids.map(getSlide);
    },
    getItem: (id: string) => getSlide(id),
    getItemAt: (index: number) => getSlide(state.ids[index]),
  };

  const context = {
    presentation: {
      slides,
      getSelectedSlides: () => ({
        load: () => undefined,
        get items() {
          return state.selectedIds.map((id) => ({ id }));
        },
      }),
      insertSlidesFromBase64: (
        _base64: string,
        options: { targetSlideId?: string },
      ) => {
        state.insertionTargets.push(options.targetSlideId);
        const insertedId = `replacement-${state.nextReplacement++}`;
        const index = options.targetSlideId
          ? state.ids.indexOf(options.targetSlideId) + 1
          : 0;
        state.ids.splice(Math.max(0, index), 0, insertedId);
      },
      setSelectedSlides: (ids: string[]) => {
        state.selectedIds = [...ids];
      },
    },
    sync: async () => {
      state.syncCount++;
      state.onSync?.(state.syncCount);
    },
  };

  return { state, context };
}

describe("withSlideZip directory binding", () => {
  it("uses the original ID, returns replacement metadata, and restores selection by ID", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"], ["b", "c"]);
    const directoryVersion = getSlideDirectoryVersion(state.ids);

    const response = await withSlideZip(
      context as never,
      {
        slide_id: "b",
        // Deliberately stale position: the ID is authoritative.
        slide_index: 0,
        directory_version: directoryVersion,
      },
      async ({ zip, markDirty }) => {
        zip.file("ppt/slides/slide1.xml", '<p:sld changed="1"/>');
        markDirty();
        return "updated";
      },
    );

    expect(state.exportedIds).toEqual(["b"]);
    expect(state.deletedIds).toEqual(["b"]);
    expect(state.insertionTargets).toEqual(["a"]);
    expect(state.ids).toEqual(["a", "replacement-1", "c"]);
    expect(state.selectedIds).toEqual(["replacement-1", "c"]);
    expect(response).toMatchObject({
      result: "updated",
      originalSlideId: "b",
      slideId: "replacement-1",
      replacementSlideId: "replacement-1",
      slideIndex: 1,
      directoryVersion: getSlideDirectoryVersion(state.ids),
    });
  });

  it("keeps the legacy numeric overload returning the callback value", async () => {
    const { state, context } = createFakeContext(["a", "b"]);
    const response = await withSlideZip(
      context as never,
      1,
      async () => "legacy",
    );

    expect(response).toBe("legacy");
    expect(state.exportedIds).toEqual(["b"]);
    expect(state.ids).toEqual(["a", "b"]);
  });

  it("rejects a stale index-only target before export or mutation", async () => {
    const originalIds = ["A", "B", "C", "D", "E", "F", "G", "H"];
    const originalVersion = getSlideDirectoryVersion(originalIds);
    const { state, context } = createFakeContext([
      "A",
      "B",
      "D",
      "E",
      "F",
      "G",
      "H",
    ]);

    await expect(
      withSlideZip(
        context as never,
        {
          slide_index: 7,
          directory_version: originalVersion,
        },
        async ({ markDirty }) => {
          markDirty();
          return null;
        },
      ),
    ).rejects.toBeInstanceOf(SlideDirectoryChangedError);

    expect(state.exportedIds).toEqual([]);
    expect(state.insertionTargets).toEqual([]);
    expect(state.deletedIds).toEqual([]);
    expect(state.ids).toEqual(["A", "B", "D", "E", "F", "G", "H"]);
  });

  it("fails closed when the directory changes before the replacement is queued", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"]);
    const directoryVersion = getSlideDirectoryVersion(state.ids);

    await expect(
      withSlideZip(
        context as never,
        {
          slide_id: "b",
          directory_version: directoryVersion,
        },
        async ({ markDirty }) => {
          state.ids.splice(0, 1);
          markDirty();
          return null;
        },
      ),
    ).rejects.toMatchObject({
      code: "SLIDE_DIRECTORY_CHANGED",
      mutationCompleted: false,
      mutationState: "not_started",
    });

    expect(state.insertionTargets).toEqual([]);
    expect(state.deletedIds).toEqual([]);
  });

  it("marks callback validation failures as not started", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"]);

    await expect(
      withSlideZip(context as never, { slide_id: "b" }, async () => {
        throw new Error("invalid callback input");
      }),
    ).rejects.toMatchObject({
      code: "SLIDE_MUTATION_NOT_STARTED",
      mutationCompleted: false,
      mutationState: "not_started",
      message: "invalid callback input",
    });

    expect(state.insertionTargets).toEqual([]);
    expect(state.deletedIds).toEqual([]);
  });

  it("marks a rejected host write sync as uncertain", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"]);
    state.onSync = (syncCount) => {
      if (syncCount === 4) throw new Error("host write sync failed");
    };

    await expect(
      withSlideZip(
        context as never,
        { slide_id: "b" },
        async ({ markDirty }) => {
          markDirty();
        },
      ),
    ).rejects.toMatchObject({
      code: "SLIDE_MUTATION_UNCERTAIN",
      mutationCompleted: true,
      mutationState: "uncertain",
      message: "host write sync failed",
    });

    expect(state.deletedIds).toEqual(["b"]);
    expect(state.ids).toContain("replacement-1");
  });

  it("refreshes selection before commit instead of restoring the start snapshot", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"], ["b"]);

    await withSlideZip(
      context as never,
      { slide_id: "b" },
      async ({ markDirty }) => {
        state.selectedIds = ["c"];
        markDirty();
      },
    );

    expect(state.selectedIds).toEqual(["c"]);
  });

  it("rechecks the lightweight directory after restoring selection", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"], ["b"]);
    state.onSync = (syncCount) => {
      if (syncCount !== 6) return;
      const moved = state.ids.pop();
      if (moved) state.ids.unshift(moved);
    };

    await expect(
      withSlideZip(
        context as never,
        { slide_id: "b" },
        async ({ markDirty }) => {
          markDirty();
        },
      ),
    ).rejects.toMatchObject({
      code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
      mutationCompleted: true,
      mutationState: "uncertain",
    });

    expect(state.deletedIds).toEqual(["b"]);
    expect(state.ids).toContain("replacement-1");
  });

  it("marks a post-write concurrent reorder as completed and non-replayable", async () => {
    const { state, context } = createFakeContext(["a", "b", "c"]);
    const directoryVersion = getSlideDirectoryVersion(state.ids);
    state.onSync = (syncCount) => {
      if (syncCount !== 4) return;
      const moved = state.ids.pop();
      if (moved) state.ids.unshift(moved);
    };

    await expect(
      withSlideZip(
        context as never,
        {
          slide_id: "b",
          directory_version: directoryVersion,
        },
        async ({ markDirty }) => {
          markDirty();
          return null;
        },
      ),
    ).rejects.toMatchObject({
      code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
      mutationCompleted: true,
      mutationState: "uncertain",
    });

    expect(state.deletedIds).toEqual(["b"]);
    expect(state.ids).toContain("replacement-1");
  });
});
