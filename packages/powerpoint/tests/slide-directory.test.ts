import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  createSlideDirectorySnapshot,
  resolveSlideTarget,
  SlideDirectoryChangedError,
} from "../src/lib/pptx/slide-directory";

const state = vi.hoisted(() => ({
  slides: [] as Array<{
    id: string;
    shapes: {
      load: () => void;
      items: Array<{
        id: string;
        type: string;
        textFrame: {
          load: () => void;
          isNullObject: boolean;
          hasText: boolean;
          textRange: { load: () => void; text: string };
        };
        getTextFrameOrNullObject: () => {
          load: () => void;
          isNullObject: boolean;
          hasText: boolean;
          textRange: { load: () => void; text: string };
        };
      }>;
    };
  }>,
  selectedIds: [] as string[],
  shapeLoads: 0,
}));

function makeSlide(id: string, texts: Array<string | null> = []) {
  return {
    id,
    shapes: {
      load: () => {
        state.shapeLoads++;
      },
      items: texts.map((text, index) => {
        const textFrame = {
          load: () => undefined,
          isNullObject: text === null,
          hasText: typeof text === "string" && text.length > 0,
          textRange: {
            load: () => undefined,
            text: text ?? "",
          },
        };
        return {
          id: String(index + 1),
          type: text === null ? "Image" : "TextBox",
          textFrame,
          getTextFrameOrNullObject: () => textFrame,
        };
      }),
    },
  };
}

const context = {
  presentation: {
    slides: {
      load: () => undefined,
      get items() {
        return state.slides;
      },
    },
    getSelectedSlides: () => ({
      load: () => undefined,
      get items() {
        return state.selectedIds.map((id) => ({ id }));
      },
    }),
  },
  sync: async () => undefined,
};

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (value: object) => unknown) => callback(context),
}));

vi.mock("../src/lib/tools/types", () => ({
  defineTool: (config: object) => config,
  toolSuccess: (data: unknown) => ({
    content: [{ type: "text", text: JSON.stringify(data) }],
    details: undefined,
  }),
  toolError: (error: string) => ({
    content: [
      { type: "text", text: JSON.stringify({ success: false, error }) },
    ],
    details: undefined,
  }),
}));

import { listSlidesTool } from "../src/lib/tools/list-slides";
import { readSlidesTool } from "../src/lib/tools/read-slides";

beforeEach(() => {
  state.slides = Array.from({ length: 30 }, (_, index) =>
    makeSlide(`slide-${index + 1}`, [`Slide ${index + 1}`]),
  );
  state.selectedIds = ["slide-28", "slide-2"];
  state.shapeLoads = 0;
});

describe("list_slides", () => {
  it("returns a content-free 25-slide directory in canonical deck order", async () => {
    const response = await listSlidesTool.execute("list-1", {});
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.items).toHaveLength(25);
    expect(payload.result.items[0]).toEqual({
      slideId: "slide-1",
      slideIndex: 0,
      positionOneIndexed: 1,
      selected: false,
    });
    expect(payload.result.items[1].selected).toBe(true);
    expect(payload.result.page).toMatchObject({
      total: 30,
      returned: 25,
      hasMore: true,
      nextOffset: 25,
    });
    expect(payload.result.ordering).toBe("deck-order");
    expect(state.shapeLoads).toBe(0);
  });

  it("paginates without repeating selected slides or relying on Slide.index", async () => {
    const response = await listSlidesTool.execute("list-2", {
      offset: 25,
      limit: 100,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(
      payload.result.items.map((item: { slideId: string }) => item.slideId),
    ).toEqual(["slide-26", "slide-27", "slide-28", "slide-29", "slide-30"]);
    expect(payload.result.items[2].selected).toBe(true);
    expect(payload.result.page).toMatchObject({
      limit: 25,
      hasMore: false,
      nextOffset: null,
    });
  });

  it("exposes a directory version mismatch when the deck changes between pages", async () => {
    const firstResponse = await listSlidesTool.execute("list-page-1", {
      offset: 0,
      limit: 25,
    });
    const firstPayload = JSON.parse(firstResponse.content[0].text);

    state.slides.splice(2, 1);

    const secondResponse = await listSlidesTool.execute("list-page-2", {
      offset: 25,
      limit: 25,
    });
    const secondPayload = JSON.parse(secondResponse.content[0].text);

    expect(firstPayload.result.directoryVersion).not.toBe(
      secondPayload.result.directoryVersion,
    );
    expect(firstPayload.result.page.total).toBe(30);
    expect(secondPayload.result.page.total).toBe(29);
    expect(
      secondPayload.result.items.map(
        (item: { slideId: string }) => item.slideId,
      ),
    ).toEqual(["slide-27", "slide-28", "slide-29", "slide-30"]);
    expect(() =>
      resolveSlideTarget(createSlideDirectorySnapshot(state.slides), {
        slide_index: 25,
        directory_version: firstPayload.result.directoryVersion,
      }),
    ).toThrowError(SlideDirectoryChangedError);

    const guardedResponse = await listSlidesTool.execute("list-page-stale", {
      offset: 25,
      limit: 25,
      directory_version: firstPayload.result.directoryVersion,
    });
    const guardedPayload = JSON.parse(guardedResponse.content[0].text);
    expect(guardedPayload.success).toBe(false);
    expect(guardedPayload.error).toContain("Slide directory changed");
  });
});

describe("read_slides", () => {
  it("reads the original target by ID after an earlier slide is deleted", async () => {
    state.slides = "ABCDEFGH"
      .split("")
      .map((id) => makeSlide(id, [`Slide ${id}`]));
    const originalDirectory = await listSlidesTool.execute("original-dir", {});
    const originalVersion = JSON.parse(originalDirectory.content[0].text).result
      .directoryVersion;

    state.slides.splice(2, 1);
    const response = await readSlidesTool.execute("read-original-h", {
      slide_ids: ["H"],
      directory_version: originalVersion,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.items[0]).toMatchObject({
      slideId: "H",
      slideIndex: 6,
      positionOneIndexed: 7,
      textPreview: "Slide H",
    });
    expect(payload.result.directoryChanged).toBe(true);
  });

  it("preserves requested ID order, removes duplicates, and truncates by Unicode code point", async () => {
    state.slides = [
      makeSlide("slide-1", ["First slide"]),
      makeSlide("slide-2", [`${"a".repeat(39)}😀tail`, null]),
    ];

    const response = await readSlidesTool.execute("read-1", {
      slide_ids: ["slide-2", "slide-1", "slide-2"],
      preview_chars: 40,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(
      payload.result.items.map((item: { slideId: string }) => item.slideId),
    ).toEqual(["slide-2", "slide-1"]);
    expect(Array.from(payload.result.items[0].textPreview)).toHaveLength(40);
    expect(payload.result.items[0].textPreview.endsWith("😀")).toBe(true);
    expect(payload.result.items[0]).toMatchObject({
      previewTruncated: true,
      shapeCount: 2,
      textFrameShapeCount: 1,
      nonTextFrameShapeCount: 1,
    });
    expect(payload.result.duplicateSlideIdsOmitted).toBe(1);
    expect(payload.result.hasMore).toBe(false);
    expect(payload.result.remainingSlideIds).toEqual([]);
    expect(payload.result.omittedFields).toContain("geometry");
    expect(payload.result.omittedFields).toContain("notes");
  });

  it("clamps direct reads to 25 slides", async () => {
    const response = await readSlidesTool.execute("read-2", {
      slide_ids: state.slides.map((slide) => slide.id),
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.items).toHaveLength(25);
    expect(payload.result.previewChars).toBe(240);
    expect(payload.result.duplicateSlideIdsOmitted).toBe(0);
    expect(payload.result.slideIdsOmittedByLimit).toBe(5);
    expect(payload.result.hasMore).toBe(true);
    expect(payload.result.remainingSlideIds).toEqual([
      "slide-26",
      "slide-27",
      "slide-28",
      "slide-29",
      "slide-30",
    ]);
  });

  it("keeps the full 25-slide emoji result below the tool output cap", async () => {
    state.slides = Array.from({ length: 25 }, (_, index) =>
      makeSlide(`emoji-${index + 1}`, ["😀".repeat(500)]),
    );

    const response = await readSlidesTool.execute("read-budget", {
      slide_ids: state.slides.map((slide) => slide.id),
      preview_chars: 500,
    });
    const bytes = new TextEncoder().encode(response.content[0].text).byteLength;
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.items).toHaveLength(25);
    expect(
      payload.result.items.every(
        (item: { previewTruncated: boolean }) => item.previewTruncated,
      ),
    ).toBe(true);
    expect(bytes).toBeLessThan(24 * 1024);
    expect(payload.result.hasMore).toBe(false);
  });

  it("budgets complete JSON items with escaped and multibyte preview text", async () => {
    state.slides = Array.from({ length: 25 }, (_, index) =>
      makeSlide(`escaped-${index + 1}`, [`中😀${'\\"'.repeat(300)}`]),
    );

    const firstResponse = await readSlidesTool.execute("read-escaped-1", {
      slide_ids: state.slides.map((slide) => slide.id),
      preview_chars: 500,
    });
    const firstText = firstResponse.content[0].text;
    const firstPayload = JSON.parse(firstText);

    expect(new TextEncoder().encode(firstText).byteLength).toBeLessThanOrEqual(
      24 * 1024,
    );
    expect(firstPayload.truncated).toBeUndefined();
    expect(firstPayload.result.items.length).toBeGreaterThan(0);
    expect(firstPayload.result.items.length).toBeLessThan(25);
    expect(firstPayload.result.items[0].textPreview).toContain('中😀\\"');
    expect(firstPayload.result.hasMore).toBe(true);
    expect(firstPayload.result.remainingSlideIds).toEqual(
      state.slides
        .slice(firstPayload.result.items.length)
        .map((slide) => slide.id),
    );
    expect(firstPayload.result.slideIdsOmittedByBudget).toBe(
      firstPayload.result.remainingSlideIds.length,
    );

    const secondResponse = await readSlidesTool.execute("read-escaped-2", {
      slide_ids: firstPayload.result.remainingSlideIds,
      preview_chars: 500,
    });
    const secondPayload = JSON.parse(secondResponse.content[0].text);

    expect(secondPayload.result.hasMore).toBe(false);
    expect(secondPayload.result.remainingSlideIds).toEqual([]);
    expect(
      [...firstPayload.result.items, ...secondPayload.result.items].map(
        (item: { slideId: string }) => item.slideId,
      ),
    ).toEqual(state.slides.map((slide) => slide.id));
  });

  it("rejects stale slide IDs and asks for a fresh directory", async () => {
    const response = await readSlidesTool.execute("read-3", {
      slide_ids: ["missing-slide"],
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain("list_slides");
  });
});
