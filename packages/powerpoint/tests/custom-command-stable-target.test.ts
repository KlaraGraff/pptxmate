import { Window } from "happy-dom";
import {
  afterAll,
  beforeAll,
  beforeEach,
  describe,
  expect,
  it,
  vi,
} from "vitest";

const mocks = vi.hoisted(() => ({
  officeContext: { kind: "powerpoint-context" },
  targets: [] as Array<Record<string, unknown>>,
}));

vi.mock("@office-agents/core", () => ({
  getSharedCustomCommands: () => ({ commands: [], promptSnippets: [] }),
}));

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (context: object) => unknown) =>
    callback(mocks.officeContext),
  withSlideZip: vi.fn(async (_context, target: Record<string, unknown>) => {
    mocks.targets.push(target);
    const usedLegacyIndex = target.slide_id === undefined;
    const slideIndex =
      typeof target.slide_index === "number" ? target.slide_index : 1;
    return {
      result: { shapeId: "42", shapeName: "inserted-shape" },
      originalSlideId: usedLegacyIndex
        ? `legacy-${slideIndex}`
        : target.slide_id,
      slideId: "replacement-42",
      replacementSlideId: "replacement-42",
      slideIndex,
      directoryVersion: "directory-v1:after",
      directoryChanged: true,
      inputDirectoryChanged: false,
      relocated: false,
      usedLegacyIndex,
    };
  }),
}));

import { withSlideZip } from "../src/lib/pptx/slide-zip";
import { getCustomCommands } from "../src/lib/vfs/custom-commands";

const pngBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0]);

async function loadCommand(name: string) {
  const result = getCustomCommands({} as never);
  const lazy = result.commands.find((command) => command.name === name);
  if (!lazy || !("load" in lazy)) throw new Error(`${name} command not found`);
  return lazy.load();
}

function commandContext() {
  return {
    cwd: "/home/user",
    fs: {
      readFileBuffer: vi.fn(async () => pngBytes),
    },
  } as never;
}

beforeAll(() => {
  const window = new Window();
  vi.stubGlobal("DOMParser", window.DOMParser);
  vi.stubGlobal(
    "Image",
    class {
      src = "";
      naturalWidth = 200;
      naturalHeight = 100;
      async decode() {}
    },
  );
  class TestURL extends URL {
    static createObjectURL() {
      return "blob:test";
    }
    static revokeObjectURL() {}
  }
  vi.stubGlobal("URL", TestURL);
  vi.stubGlobal("document", {
    createElement: (tag: string) => {
      if (tag !== "canvas") throw new Error(`Unexpected element: ${tag}`);
      return {
        width: 0,
        height: 0,
        getContext: () => ({ drawImage: () => undefined }),
        toBlob: (callback: (blob: Blob) => void) =>
          callback(new Blob([pngBytes], { type: "image/png" })),
      };
    },
  });
});

afterAll(() => {
  vi.unstubAllGlobals();
});

beforeEach(() => {
  mocks.targets.length = 0;
  vi.stubGlobal(
    "fetch",
    vi.fn(async () => ({
      ok: true,
      status: 200,
      statusText: "OK",
      text: async () =>
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"></svg>',
    })),
  );
});

describe("PowerPoint insert commands", () => {
  it("targets insert-image by stable slide ID and returns a mutation receipt", async () => {
    const command = await loadCommand("insert-image");
    const result = await command.execute(
      [
        "/home/user/image.png",
        "--slide-id=slide-stable",
        "--directory-version=directory-v1:before",
      ],
      commandContext(),
    );
    const receipt = JSON.parse(result.stdout);

    expect(result).toMatchObject({ exitCode: 0, stderr: "" });
    expect(mocks.targets).toEqual([
      {
        slide_id: "slide-stable",
        directory_version: "directory-v1:before",
      },
    ]);
    expect(receipt).toMatchObject({
      success: true,
      operation: "insert-image",
      slideId: "replacement-42",
      replacementSlideId: "replacement-42",
      directoryVersion: "directory-v1:after",
      mutationCompleted: true,
      mutationState: "completed",
      usedLegacyIndex: false,
    });
  });

  it("keeps the legacy 1-based slide number as an object target", async () => {
    const command = await loadCommand("insert-image");
    const result = await command.execute(
      ["/home/user/image.png", "2"],
      commandContext(),
    );
    const receipt = JSON.parse(result.stdout);

    expect(mocks.targets).toEqual([{ slide_index: 1 }]);
    expect(receipt).toMatchObject({
      success: true,
      slideIndex: 1,
      positionOneIndexed: 2,
      usedLegacyIndex: true,
    });
  });

  it.each([
    ["in", 5, 2.5],
    ["cm", 12.7, 6.35],
    ["emu", 4_572_000, 2_286_000],
  ])("keeps the default image size physical when using %s units", async (unit, expectedWidth, expectedHeight) => {
    const command = await loadCommand("insert-image");
    const result = await command.execute(
      ["/home/user/image.png", "--slide-id=slide-stable", `--unit=${unit}`],
      commandContext(),
    );
    const receipt = JSON.parse(result.stdout);

    expect(result.exitCode).toBe(0);
    expect(receipt.unit).toBe(unit);
    expect(receipt.width).toBeCloseTo(expectedWidth);
    expect(receipt.height).toBeCloseTo(expectedHeight);
  });

  it("accepts stable IDs for insert-icon", async () => {
    const command = await loadCommand("insert-icon");
    const result = await command.execute(
      ["mdi:alert", "--slide-id=slide-icon", "--directory-version=dir-icon"],
      commandContext(),
    );
    const receipt = JSON.parse(result.stdout);

    expect(result.exitCode).toBe(0);
    expect(mocks.targets).toEqual([
      { slide_id: "slide-icon", directory_version: "dir-icon" },
    ]);
    expect(receipt).toMatchObject({
      success: true,
      operation: "insert-icon",
      replacementSlideId: "replacement-42",
      directoryVersion: "directory-v1:after",
      mutationCompleted: true,
      mutationState: "completed",
    });
  });

  it("returns structured JSON on stdout for preflight failures", async () => {
    const command = await loadCommand("insert-image");
    const result = await command.execute(
      ["/home/user/image.png"],
      commandContext(),
    );
    const receipt = JSON.parse(result.stdout);

    expect(result).toMatchObject({ exitCode: 1, stderr: "" });
    expect(receipt).toMatchObject({
      success: false,
      operation: "insert-image",
      mutationCompleted: false,
      mutationState: "not_started",
      error: { code: "SLIDE_TARGET_REQUIRED" },
    });
  });

  it("marks post-write directory failures as uncertain at the top level", async () => {
    vi.mocked(withSlideZip).mockRejectedValueOnce(
      Object.assign(new Error("directory changed during replacement"), {
        code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
        mutationCompleted: true,
        expectedVersion: "directory-v1:before",
        currentVersion: "directory-v1:concurrent",
      }),
    );
    const command = await loadCommand("insert-image");
    const result = await command.execute(
      ["/home/user/image.png", "--slide-id=slide-stable"],
      commandContext(),
    );
    const receipt = JSON.parse(result.stdout);

    expect(result).toMatchObject({ exitCode: 1, stderr: "" });
    expect(receipt).toMatchObject({
      success: false,
      operation: "insert-image",
      mutationCompleted: true,
      mutationState: "uncertain",
      expectedDirectoryVersion: "directory-v1:before",
      currentDirectoryVersion: "directory-v1:concurrent",
      error: { code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE" },
    });
  });
});
