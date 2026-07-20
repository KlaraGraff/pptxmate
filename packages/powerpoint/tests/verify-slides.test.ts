import { describe, expect, it, vi } from "vitest";

const { context } = vi.hoisted(() => {
  const makeSlide = (id: string) => ({
    id: `slide-${id}`,
    shapes: {
      load: () => undefined,
      items: [
        {
          id,
          name: `Shape ${id}`,
          left: 95,
          top: 10,
          width: 10,
          height: 10,
        },
      ],
    },
  });
  return {
    context: {
      presentation: {
        slides: {
          load: () => undefined,
          items: [makeSlide("1"), makeSlide("2")],
          getItem(id: string) {
            return this.items.find((slide) => slide.id === id);
          },
        },
        pageSetup: {
          load: () => undefined,
          slideWidth: 100,
          slideHeight: 100,
        },
      },
      sync: async () => undefined,
    },
  };
});

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

import { verifySlidesTool } from "../src/lib/tools/verify-slides";

describe("verify_slides limits", () => {
  it("marks the result truncated when later slides were not checked", async () => {
    const response = await verifySlidesTool.execute("call-1", {
      max_issues: 1,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.issues).toHaveLength(1);
    expect(payload.result.issues[0].slideId).toBe("slide-1");
    expect(payload.result.checkedSlideCount).toBe(2);
    expect(payload.result.issueCount).toBe(2);
    expect(payload.result.truncated).toBe(true);
  });

  it("rejects negative slide indices", async () => {
    const response = await verifySlidesTool.execute("call-2", {
      slide_indices: [-1],
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain("out of range");
  });
});
