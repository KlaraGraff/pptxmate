import { Window } from "happy-dom";
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { MAX_PLAIN_TEXT_WRITE_BYTES } from "../src/lib/pptx/text-xml";

const state = vi.hoisted(() => ({
  slideXml: "",
  writes: 0,
  dirtyCalls: 0,
}));

function makeSlideXml(text: string): string {
  return `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Long Text"/></p:nvSpPr>
          <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:r><a:rPr b="1"/><a:t>${text}</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
      </p:spTree></p:cSld>
    </p:sld>`;
}

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (context: object) => unknown) => callback({}),
  withSlideZip: async (
    _context: object,
    _slideIndex: number,
    callback: (value: object) => unknown,
  ) =>
    callback({
      zip: {
        file: (_path: string, content?: string) => {
          if (content !== undefined) {
            state.slideXml = content;
            state.writes++;
            return undefined;
          }
          return { async: async () => state.slideXml };
        },
      },
      markDirty: () => {
        state.dirtyCalls++;
      },
    }),
}));

vi.mock("../src/lib/tools/types", () => ({
  defineTool: (config: object) => config,
  toolSuccess: (data: unknown) => ({
    content: [{ type: "text", text: JSON.stringify(data) }],
    details: undefined,
  }),
  toolError: (error: string, metadata?: Record<string, unknown>) => ({
    content: [
      {
        type: "text",
        text: JSON.stringify({ success: false, error, ...metadata }),
      },
    ],
    details: undefined,
  }),
}));

import { editSlideTextTool } from "../src/lib/tools/edit-slide-text";
import { readSlideTextTool } from "../src/lib/tools/read-slide-text";
import { updateSlideTextTool } from "../src/lib/tools/update-slide-text";

beforeAll(() => {
  const window = new Window();
  const getElementsByTagNameNS = function (
    this: Element | Document,
    namespace: string,
    localName: string,
  ) {
    return Array.from(this.getElementsByTagName("*")).filter(
      (element) =>
        element.namespaceURI === namespace && element.localName === localName,
    ) as unknown as HTMLCollectionOf<Element>;
  };
  window.Element.prototype.getElementsByTagNameNS = getElementsByTagNameNS;
  window.Document.prototype.getElementsByTagNameNS = getElementsByTagNameNS;
  window.XMLDocument.prototype.getElementsByTagNameNS = getElementsByTagNameNS;
  vi.stubGlobal("DOMParser", window.DOMParser);
  vi.stubGlobal("XMLSerializer", window.XMLSerializer);
});

beforeEach(() => {
  state.slideXml = makeSlideXml(`${"A".repeat(1_500)}TAIL`);
  state.writes = 0;
  state.dirtyCalls = 0;
});

describe("paged plain-text writeback", () => {
  it("continues from an adjusted cursor after a longer replacement", async () => {
    const firstRead = await readSlideTextTool.execute("read-1", {
      slide_index: 0,
      shape_id: "2",
      max_bytes: 1_000,
    });
    const firstPayload = JSON.parse(firstRead.content[0].text);
    const item = firstPayload.result.paragraphs[0];

    expect(item.text).toBe("A".repeat(1_000));
    expect(item.editScope).toEqual({
      paragraph_start: 0,
      paragraph_end: 1,
      char_start: 0,
      char_end: 1_000,
    });

    const replacement = "B".repeat(1_200);
    const edit = await editSlideTextTool.execute("edit-1", {
      slide_index: 0,
      shape_id: "2",
      text: replacement,
      ...item.editScope,
      expected_text_hash: item.textHash,
    });
    const editPayload = JSON.parse(edit.content[0].text);

    expect(editPayload.success).toBe(true);
    expect(editPayload).toMatchObject({ slideIndex: 0, shapeId: "2" });
    expect(editPayload.nextCursor).toEqual({
      paragraph_offset: 0,
      char_offset: 1_200,
    });
    expect(state.writes).toBe(1);
    expect(state.dirtyCalls).toBe(1);

    const secondRead = await readSlideTextTool.execute("read-2", {
      slide_index: 0,
      shape_id: "2",
      ...editPayload.nextCursor,
      max_bytes: 1_000,
    });
    const secondPayload = JSON.parse(secondRead.content[0].text);

    expect(secondPayload.result.paragraphs[0]).toMatchObject({
      text: `${"A".repeat(500)}TAIL`,
      editScope: {
        paragraph_start: 0,
        paragraph_end: 1,
        char_start: 1_200,
        char_end: 1_704,
      },
    });
  });

  it("does not write when the expected hash is stale", async () => {
    const response = await editSlideTextTool.execute("edit-stale", {
      slide_index: 0,
      shape_id: "2",
      text: "replacement",
      paragraph_start: 0,
      paragraph_end: 1,
      char_start: 0,
      char_end: 10,
      expected_text_hash: "fnv1a32:00000000",
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain("Text guard mismatch");
    expect(state.writes).toBe(0);
    expect(state.dirtyCalls).toBe(0);
  });

  it("accepts an exact original-text guard as an alternative to a hash", async () => {
    const original = `${"A".repeat(1_500)}TAIL`;
    const response = await editSlideTextTool.execute("edit-exact", {
      slide_index: 0,
      shape_id: "2",
      text: "translated",
      paragraph_start: 0,
      paragraph_end: 1,
      expected_text: original,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(true);
    expect(payload.beforeTextHash).toMatch(/^fnv1a32:/);
    expect(state.slideXml).toContain("translated");
    expect(state.writes).toBe(1);
  });

  it("requires a guard for every range write", async () => {
    const response = await editSlideTextTool.execute("edit-unguarded", {
      slide_index: 0,
      shape_id: "2",
      text: "replacement",
      paragraph_start: 0,
      paragraph_end: 1,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain("require expected_text_hash");
    expect(payload).toMatchObject({
      mutationCompleted: false,
      mutationState: "not_started",
    });
    expect(state.writes).toBe(0);
  });

  it("rejects plain-text payloads above the UTF-8 write limit", async () => {
    const response = await editSlideTextTool.execute("edit-oversized", {
      slide_index: 0,
      shape_id: "2",
      text: "x".repeat(MAX_PLAIN_TEXT_WRITE_BYTES + 1),
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain(
      `per-call limit is ${MAX_PLAIN_TEXT_WRITE_BYTES}`,
    );
    expect(state.writes).toBe(0);
  });

  it("keeps the legacy whole-shape replacement path", async () => {
    const response = await editSlideTextTool.execute("edit-legacy", {
      slide_index: 0,
      shape_id: "2",
      text: "Short replacement",
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload).toEqual({ success: true });
    expect(state.slideXml).toContain("Short replacement");
    expect(state.writes).toBe(1);
  });

  it("validates a whole-shape hash before a batch update", async () => {
    const read = await readSlideTextTool.execute("read-batch", {
      slide_index: 0,
      shape_id: "2",
    });
    const shapeTextHash = JSON.parse(read.content[0].text).result.shapeTextHash;
    const response = await updateSlideTextTool.execute("update-guarded", {
      slide_index: 0,
      updates: [
        {
          shape_id: "2",
          text: "batch replacement",
          expected_text_hash: shapeTextHash,
        },
      ],
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(true);
    expect(state.slideXml).toContain("batch replacement");
    expect(state.writes).toBe(1);
  });

  it("rejects a batch above its aggregate write limit", async () => {
    const response = await updateSlideTextTool.execute("update-oversized", {
      slide_index: 0,
      updates: Array.from({ length: 5 }, () => ({
        shape_id: "2",
        text: "x".repeat(MAX_PLAIN_TEXT_WRITE_BYTES),
      })),
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain("Batch text");
    expect(state.writes).toBe(0);
  });

  it("counts exact original-text guards against the batch limit", async () => {
    const response = await updateSlideTextTool.execute(
      "update-oversized-guards",
      {
        slide_index: 0,
        updates: Array.from({ length: 5 }, () => ({
          shape_id: "2",
          text: "x",
          expected_text: "A".repeat(MAX_PLAIN_TEXT_WRITE_BYTES),
        })),
      },
    );
    const payload = JSON.parse(response.content[0].text);

    expect(payload.success).toBe(false);
    expect(payload.error).toContain("use expected_text_hash");
    expect(state.writes).toBe(0);
  });
});
