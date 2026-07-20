import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  officeContext: { kind: "powerpoint-context" },
  cleanup: vi.fn(),
  withSlideZip: vi.fn(),
}));

vi.mock("@office-agents/core", () => ({
  sandboxedEval: vi.fn(async () => ({ edited: true })),
}));

vi.mock("../src/lib/pptx/master-cleanup", () => ({
  cleanupSlideMasters: mocks.cleanup,
}));

vi.mock("../src/lib/pptx/slide-directory", () => ({
  loadSlideDirectory: vi.fn(async () => ({
    slideIds: ["slide-original"],
    directoryVersion: "directory-v1:before",
    indexById: new Map([["slide-original", 0]]),
  })),
}));

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (context: object) => unknown) =>
    callback(mocks.officeContext),
  withSlideZip: mocks.withSlideZip,
}));

vi.mock("../src/lib/tools/types", () => ({
  defineTool: (config: object) => config,
  toolSuccess: (data: unknown) => ({
    content: [{ type: "text", text: JSON.stringify(data) }],
    details: undefined,
  }),
}));

import { createEditSlideMasterTool } from "../src/lib/tools/edit-slide-master";

const zipReceipt = {
  result: { edited: true },
  originalSlideId: "slide-original",
  slideId: "slide-replacement",
  replacementSlideId: "slide-replacement",
  slideIndex: 0,
  directoryVersion: "directory-v1:after",
  directoryChanged: true,
  inputDirectoryChanged: false,
  relocated: false,
  usedLegacyIndex: false,
};

beforeEach(() => {
  mocks.withSlideZip.mockReset();
  mocks.withSlideZip.mockResolvedValue({ ...zipReceipt });
  mocks.cleanup.mockReset();
  mocks.cleanup.mockResolvedValue({
    directoryVersion: "directory-v1:after",
    reassignedSlideCount: 1,
    orphanedMasterCount: 1,
    temporarySlidesCreated: 1,
  });
});

describe("edit_slide_master receipts", () => {
  it("returns replacement and cleanup mutation metadata", async () => {
    const tool = createEditSlideMasterTool({} as never);
    const response = await tool.execute("master-1", { code: "markDirty();" });
    const payload = JSON.parse(response.content[0].text);

    expect(mocks.withSlideZip).toHaveBeenCalledWith(
      mocks.officeContext,
      {
        slide_id: "slide-original",
        directory_version: "directory-v1:before",
      },
      expect.any(Function),
    );
    expect(mocks.cleanup).toHaveBeenCalledWith(
      mocks.officeContext,
      "directory-v1:after",
    );
    expect(payload).toMatchObject({
      success: true,
      slideId: "slide-replacement",
      replacementSlideId: "slide-replacement",
      directoryVersion: "directory-v1:after",
      mutationCompleted: true,
      mutationState: "completed",
      mutation: {
        kind: "master",
        status: "completed",
        cleanup: { temporarySlidesCreated: 1 },
      },
    });
  });

  it("preserves the completed main write when cleanup becomes uncertain", async () => {
    mocks.cleanup.mockRejectedValue(
      Object.assign(new Error("directory changed during cleanup"), {
        code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
        mutationCompleted: true,
        cleanupPhase: "before_delete_temporary_slide",
        expectedVersion: "directory-v1:temp",
        currentVersion: "directory-v1:concurrent",
        temporarySlideId: "temporary-slide",
      }),
    );
    const tool = createEditSlideMasterTool({} as never);
    const response = await tool.execute("master-2", { code: "markDirty();" });
    const payload = JSON.parse(response.content[0].text);

    expect(payload).toMatchObject({
      success: false,
      slideId: "slide-replacement",
      replacementSlideId: "slide-replacement",
      directoryVersion: "directory-v1:after",
      mutationCompleted: true,
      mutationState: "uncertain",
      cleanupPhase: "before_delete_temporary_slide",
      expectedDirectoryVersion: "directory-v1:temp",
      currentDirectoryVersion: "directory-v1:concurrent",
      temporarySlideId: "temporary-slide",
      error: { code: "SLIDE_DIRECTORY_CHANGED_DURING_WRITE" },
    });
  });
});
