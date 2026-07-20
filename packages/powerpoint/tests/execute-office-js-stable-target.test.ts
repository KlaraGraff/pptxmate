import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  sandboxedEval: vi.fn(),
  slides: [] as Array<{ id: string }>,
  getItem: vi.fn(),
  load: vi.fn(),
  sync: vi.fn(async () => undefined),
}));

vi.mock("@office-agents/core", () => ({
  resizeImage: vi.fn(),
  sandboxedEval: mocks.sandboxedEval,
}));

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (context: object) => unknown) =>
    callback({
      presentation: {
        slides: {
          get items() {
            return mocks.slides;
          },
          load: mocks.load,
          getItem: mocks.getItem,
        },
      },
      sync: mocks.sync,
    }),
}));

import { getSlideDirectoryVersion } from "../src/lib/pptx/slide-directory";
import { createExecuteOfficeJsTool } from "../src/lib/tools/execute-office-js";

const ctx = {
  readFile: vi.fn(),
  readFileBuffer: vi.fn(),
  writeFile: vi.fn(),
};

const tool = createExecuteOfficeJsTool(ctx as never);

function parseResult(result: Awaited<ReturnType<(typeof tool)["execute"]>>) {
  const first = result.content[0];
  if (!first || first.type !== "text") throw new Error("Missing text result");
  return JSON.parse(first.text) as Record<string, unknown>;
}

beforeEach(() => {
  vi.clearAllMocks();
  mocks.slides = "ABDEFGH".split("").map((id) => ({ id }));
  mocks.getItem.mockImplementation((id: string) => ({ id, shapes: {} }));
  mocks.sandboxedEval.mockResolvedValue({ changed: true });
  vi.stubGlobal("PowerPoint", {});
  vi.stubGlobal("Office", {});
});

describe("execute_office_js stable slide target", () => {
  it("keeps the original H target after an earlier slide is deleted", async () => {
    const originalVersion = getSlideDirectoryVersion("ABCDEFGH".split(""));

    const result = await tool.execute("stable-target", {
      slide_id: "H",
      slide_index: 7,
      directory_version: originalVersion,
      code: "return { changed: true };",
    });
    const payload = parseResult(result);

    expect(payload).toMatchObject({
      success: true,
      slideId: "H",
      slideIndex: 6,
      positionOneIndexed: 7,
      directoryChanged: true,
      relocated: true,
      usedLegacyIndex: false,
      mutationCompleted: true,
      mutationState: "completed",
    });
    expect(mocks.getItem).toHaveBeenCalledWith("H");
    expect(mocks.sandboxedEval).toHaveBeenCalledTimes(1);
    const globals = mocks.sandboxedEval.mock.calls[0][1];
    expect(globals).toMatchObject({
      targetSlide: { id: "H" },
      targetSlideId: "H",
      targetSlideIndex: 6,
      directoryVersion: getSlideDirectoryVersion("ABDEFGH".split("")),
    });
  });

  it("rejects an index-only target from a stale directory before execution", async () => {
    const result = await tool.execute("stale-index", {
      slide_index: 7,
      directory_version: getSlideDirectoryVersion("ABCDEFGH".split("")),
      code: "return true;",
    });
    const payload = parseResult(result);

    expect(payload).toMatchObject({
      success: false,
      code: "SLIDE_DIRECTORY_CHANGED",
      mutationCompleted: false,
      mutationState: "not_started",
    });
    expect(mocks.sandboxedEval).not.toHaveBeenCalled();
  });

  it("rejects direct positional slide access before execution", async () => {
    const result = await tool.execute("positional-target", {
      code: "const slide = context.presentation.slides.getItemAt(7); return slide;",
    });
    const payload = parseResult(result);

    expect(payload).toMatchObject({
      success: false,
      code: "POSITIONAL_SLIDE_ACCESS_REJECTED",
      mutationCompleted: false,
      mutationState: "not_started",
    });
    expect(mocks.sync).not.toHaveBeenCalled();
    expect(mocks.sandboxedEval).not.toHaveBeenCalled();
  });

  it("requires a target argument when code uses targetSlide", async () => {
    const result = await tool.execute("missing-target", {
      code: "targetSlide.shapes.addTextBox('x');",
    });
    const payload = parseResult(result);

    expect(payload).toMatchObject({
      success: false,
      code: "SLIDE_TARGET_REQUIRED",
      mutationCompleted: false,
      mutationState: "not_started",
    });
    expect(mocks.sandboxedEval).not.toHaveBeenCalled();
  });

  it("marks failures after entering user code as uncertain", async () => {
    mocks.sandboxedEval.mockRejectedValueOnce(new Error("sync failed"));
    const result = await tool.execute("uncertain-target", {
      slide_id: "H",
      code: "targetSlide.shapes.addTextBox('x'); await context.sync();",
    });
    const payload = parseResult(result);

    expect(payload).toMatchObject({
      success: false,
      slideId: "H",
      slideIndex: 6,
      mutationCompleted: true,
      mutationState: "uncertain",
    });
  });
});
