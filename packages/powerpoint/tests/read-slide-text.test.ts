import { Window } from "happy-dom";
import { beforeAll, describe, expect, it, vi } from "vitest";

const { slideXml } = vi.hoisted(() => {
  const manyParagraphs = Array.from(
    { length: 200 },
    (_, index) =>
      `<a:p><a:r><a:t>Paragraph P${index + 1} short text</a:t></a:r></a:p>`,
  ).join("");
  return {
    slideXml: `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Text Box 1"/></p:nvSpPr>
          <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:r><a:t>First</a:t></a:r></a:p>
            <a:p><a:r><a:t>Second</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Long Text"/></p:nvSpPr>
          <p:txBody><a:bodyPr/><a:lstStyle/>${manyParagraphs}</p:txBody>
        </p:sp>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl/></a:graphicData></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`,
  };
});

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (context: object) => unknown) => callback({}),
  withSlideZip: async (
    _context: object,
    _slideIndex: number,
    callback: (value: object) => unknown,
  ) =>
    callback({
      zip: {
        file: () => ({ async: async () => slideXml }),
      },
    }),
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

import { readSlideTextTool } from "../src/lib/tools/read-slide-text";
import { readSlideTextsTool } from "../src/lib/tools/read-slide-texts";

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

describe("read_slide_text plain pagination", () => {
  it("defaults to plain text when format is omitted", async () => {
    const response = await readSlideTextTool.execute("call-default", {
      slide_index: 0,
      shape_id: "2",
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result).toMatchObject({
      format: "plain",
      slideIndex: 0,
      shapeId: "2",
    });
    expect(payload.result.paragraphs).toMatchObject([
      {
        index: 0,
        text: "First",
        editScope: { paragraph_start: 0, paragraph_end: 1 },
      },
      {
        index: 1,
        text: "Second",
        editScope: { paragraph_start: 1, paragraph_end: 2 },
      },
    ]);
    expect(payload.result.paragraphs[0].textHash).toMatch(/^fnv1a32:/);
    expect(JSON.stringify(payload)).not.toContain("<a:p>");
  });

  it("returns a null cursor after the last paragraph", async () => {
    const response = await readSlideTextTool.execute("call-1", {
      slide_index: 0,
      shape_id: "2",
      format: "plain",
    });
    const payload = JSON.parse(response.content[0].text);

    expect(
      payload.result.paragraphs.map((item: { text: string }) => item.text),
    ).toEqual(["First", "Second"]);
    expect(payload.result.page).toMatchObject({
      hasMore: false,
      nextOffset: null,
      nextCharOffset: null,
    });
  });

  it("returns the next paragraph cursor when the page is limited", async () => {
    const response = await readSlideTextTool.execute("call-2", {
      slide_index: 0,
      shape_id: "2",
      format: "plain",
      paragraph_limit: 1,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.page).toMatchObject({
      hasMore: true,
      nextOffset: 1,
      nextCharOffset: null,
      nextCursor: { paragraph_offset: 1 },
      editScope: { paragraph_start: 0, paragraph_end: 1 },
    });
    expect(payload.result.page.textHash).toBe(
      payload.result.paragraphs[0].textHash,
    );
  });

  it("keeps 200 short paragraphs within the final structured-result budget", async () => {
    const response = await readSlideTextTool.execute("call-budget", {
      slide_index: 0,
      shape_id: "3",
      format: "plain",
      paragraph_limit: 200,
      max_bytes: 8_000,
    });
    const serialized = response.content[0].text;
    const payload = JSON.parse(serialized);

    expect(new TextEncoder().encode(serialized).byteLength).toBeLessThanOrEqual(
      24 * 1024,
    );
    expect(payload.resultPreview).toBeUndefined();
    expect(payload.result.paragraphCount).toBe(200);
    expect(payload.result.paragraphs.length).toBeGreaterThan(0);
    expect(payload.result.paragraphs.length).toBeLessThan(200);
    expect(payload.result.page).toMatchObject({
      hasMore: true,
      budgetLimited: true,
    });
    expect(payload.result.page.editScope).toEqual({
      paragraph_start: 0,
      paragraph_end: payload.result.paragraphs.length,
    });
    expect(payload.result.page.nextCursor).toEqual({
      paragraph_offset: payload.result.paragraphs.length,
    });
  });

  it("returns OOXML pagination metadata instead of silently dropping later paragraphs", async () => {
    const response = await readSlideTextTool.execute("call-3", {
      slide_index: 0,
      shape_id: "3",
      format: "ooxml",
      paragraph_limit: 20,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.xml).toContain("P20");
    expect(payload.result.paragraphCount).toBe(200);
    expect(payload.result.page).toMatchObject({
      returned: 20,
      hasMore: true,
      nextOffset: 20,
    });
  });

  it("reports containers whose text is outside the compact text-box path", async () => {
    const response = await readSlideTextsTool.execute("call-4", {
      slide_index: 0,
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload.result.omittedShapeCounts).toEqual({ table: 1 });
    expect(payload.result.omittedShapes).toContainEqual({
      id: "4",
      name: "Table 1",
      type: "table",
    });
  });
});
