import { describe, expect, it } from "vitest";
import {
  createSlideDirectorySnapshot,
  getSlideDirectoryVersion,
  resolveSlideTarget,
  SlideDirectoryChangedError,
  SlideTargetNotFoundError,
} from "../src/lib/pptx/slide-directory";
import { getSlideDirectoryVersion as getPublicDirectoryVersion } from "../src/lib/slide-directory";

function snapshot(ids: string[]) {
  return createSlideDirectorySnapshot(ids.map((id) => ({ id })));
}

describe("slide directory helper", () => {
  it("fingerprints membership and order, but not selection or content", () => {
    expect(getSlideDirectoryVersion(["a", "b"])).toBe(
      getSlideDirectoryVersion(["a", "b"]),
    );
    expect(getSlideDirectoryVersion(["a", "b"])).not.toBe(
      getSlideDirectoryVersion(["b", "a"]),
    );
    expect(getSlideDirectoryVersion(["a", "b"])).not.toBe(
      getSlideDirectoryVersion(["a", "b", "c"]),
    );
    expect(getSlideDirectoryVersion(["a", "b"])).toBe(
      getPublicDirectoryVersion(["a", "b"]),
    );
  });

  it("targets original H by ID after C is deleted and H moves to index 6", () => {
    const original = snapshot(["A", "B", "C", "D", "E", "F", "G", "H"]);
    const current = snapshot(["A", "B", "D", "E", "F", "G", "H"]);
    const target = resolveSlideTarget(current, {
      slide_id: "H",
      slide_index: 7,
      directory_version: original.directoryVersion,
    });

    expect(target).toMatchObject({
      slideId: "H",
      slideIndex: 6,
      directoryVersion: current.directoryVersion,
      directoryChanged: true,
      indexMismatch: true,
      usedLegacyIndex: false,
    });
  });

  it("rejects an index-only target when its directory version is stale", () => {
    const original = snapshot(["A", "B", "C", "D", "E", "F", "G", "H"]);
    const current = snapshot(["A", "B", "D", "E", "F", "G", "H"]);
    expect(() =>
      resolveSlideTarget(current, {
        slide_index: 7,
        directory_version: original.directoryVersion,
      }),
    ).toThrowError(SlideDirectoryChangedError);
  });

  it("does not fall back to an old index when the stable target was deleted", () => {
    const current = snapshot(["A", "B", "C", "D", "E", "F", "G"]);
    expect(() =>
      resolveSlideTarget(current, {
        slide_id: "H",
        slide_index: 7,
      }),
    ).toThrowError(SlideTargetNotFoundError);
  });

  it("resolves a reordered target by ID and reports its current index", () => {
    const original = snapshot(["A", "B", "C", "D", "E", "F", "G", "H"]);
    const current = snapshot(["H", "A", "B", "C", "D", "E", "F", "G"]);

    expect(
      resolveSlideTarget(current, {
        slide_id: "H",
        slide_index: 7,
        directory_version: original.directoryVersion,
      }),
    ).toMatchObject({
      slideId: "H",
      slideIndex: 0,
      directoryChanged: true,
      indexMismatch: true,
      usedLegacyIndex: false,
    });
  });
});
