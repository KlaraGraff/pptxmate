import { Window } from "happy-dom";
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import { hashPlainText } from "../src/lib/pptx/text-xml";

const state = vi.hoisted(() => ({
  slideXml: "",
  writes: 0,
}));

const RELATIONSHIPS_XML = `
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  </Relationships>`;

const CHART_XML = `
  <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Chart title</a:t></a:r></a:p></c:rich></c:tx></c:title>
      <c:plotArea><c:barChart><c:ser><c:cat><c:strLit><c:pt idx="0"><c:v>Category one</c:v></c:pt></c:strLit></c:cat></c:ser></c:barChart></c:plotArea>
    </c:chart>
  </c:chartSpace>`;

function makeSlideXml(): string {
  return `
    <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:cSld><p:spTree>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="2" name="Long source"/></p:nvSpPr>
          <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${"A".repeat(1_600)}</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:sp>
          <p:nvSpPr><p:cNvPr id="3" name="Bilingual text"/></p:nvSpPr>
          <p:txBody><a:bodyPr/><a:lstStyle/>
            <a:p><a:r><a:t>Original source</a:t></a:r></a:p>
            <a:p><a:r><a:t>Existing translation</a:t></a:r></a:p>
          </p:txBody>
        </p:sp>
        <p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="10" name="Group 1"/></p:nvGrpSpPr>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="11" name="Grouped text"/></p:nvSpPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Grouped source</a:t></a:r></a:p></p:txBody>
          </p:sp>
        </p:grpSp>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
            <a:tr><a:tc><a:txBody><a:p><a:r><a:t>Table source</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
          </a:tbl></a:graphicData></a:graphic>
        </p:graphicFrame>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId7"/></a:graphicData></a:graphic>
        </p:graphicFrame>
        <p:graphicFrame>
          <p:nvGraphicFramePr><p:cNvPr id="6" name="Diagram 1"/></p:nvGraphicFramePr>
          <a:graphic><a:graphicData uri="urn:example:diagram"/></a:graphic>
        </p:graphicFrame>
      </p:spTree></p:cSld>
    </p:sld>`;
}

vi.mock("../src/lib/pptx/slide-zip", () => ({
  safeRun: async (callback: (context: object) => unknown) => callback({}),
  withSlideZip: async (
    _context: object,
    _target: object,
    callback: (value: object) => unknown,
  ) =>
    callback({
      zip: {
        file: (path: string, content?: string) => {
          if (content !== undefined) {
            if (path === "ppt/slides/slide1.xml") {
              state.slideXml = content;
              state.writes++;
            }
            return undefined;
          }
          const files: Record<string, string> = {
            "ppt/slides/slide1.xml": state.slideXml,
            "ppt/slides/_rels/slide1.xml.rels": RELATIONSHIPS_XML,
            "ppt/charts/chart1.xml": CHART_XML,
          };
          const value = files[path];
          return value === undefined ? null : { async: async () => value };
        },
      },
      markDirty: () => {},
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

import { patchSlideTextTool } from "../src/lib/tools/patch-slide-text";
import { readSlideTranslatableTextsTool } from "../src/lib/tools/read-slide-translatable-texts";

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
  state.slideXml = makeSlideXml();
  state.writes = 0;
});

describe("translation audit tools", () => {
  it("reads grouped shapes, table cells, and chart XML while reporting unsupported containers", async () => {
    const response = await readSlideTranslatableTextsTool.execute("audit-all", {
      slide_index: 0,
      max_bytes: 8_000,
    });
    const payload = JSON.parse(response.content[0].text);
    const result = payload.result;

    expect(result.items.map((item: { kind: string }) => item.kind)).toEqual([
      "shape",
      "shape",
      "groupShape",
      "tableCell",
      "chart",
    ]);
    expect(
      result.items.find((item: { kind: string }) => item.kind === "groupShape"),
    ).toMatchObject({
      text: "Grouped source",
      location: { shapeId: "11", groupShapeIds: ["10"] },
    });
    expect(
      result.items.find((item: { kind: string }) => item.kind === "tableCell"),
    ).toMatchObject({
      text: "Table source",
      location: { shapeId: "4", rowIndex: 0, columnIndex: 0 },
    });
    expect(
      result.items.find((item: { kind: string }) => item.kind === "chart"),
    ).toMatchObject({
      text: "Chart title\nCategory one",
      location: { shapeId: "5", chartPath: "ppt/charts/chart1.xml" },
    });
    expect(result.page).toMatchObject({ hasMore: false, nextCursor: null });
    expect(result.coverage).toMatchObject({
      readableSourceCounts: {
        shape: 2,
        groupShape: 1,
        tableCell: 1,
        chart: 1,
      },
      unsupportedContainerCount: 1,
      scanComplete: false,
    });
  });

  it("continues a long content source with the returned character cursor", async () => {
    const first = await readSlideTranslatableTextsTool.execute("audit-page-1", {
      slide_index: 0,
      max_bytes: 1_000,
    });
    const firstPayload = JSON.parse(first.content[0].text).result;

    expect(firstPayload.items).toHaveLength(1);
    expect(firstPayload.items[0]).toMatchObject({
      sourceIndex: 0,
      kind: "shape",
      textTruncated: true,
      span: { charStart: 0 },
    });
    expect(firstPayload.page).toMatchObject({
      hasMore: true,
      nextCursor: { offset: 0 },
    });
    expect(firstPayload.page.nextCursor.char_offset).toBeGreaterThan(0);

    const second = await readSlideTranslatableTextsTool.execute(
      "audit-page-2",
      {
        slide_index: 0,
        max_bytes: 1_000,
        ...firstPayload.page.nextCursor,
      },
    );
    const secondPayload = JSON.parse(second.content[0].text).result;

    expect(secondPayload.items[0]).toMatchObject({
      sourceIndex: 0,
      span: { charStart: firstPayload.page.nextCursor.char_offset },
    });
  });

  it("patches only the selected translation paragraph and returns a verification scope", async () => {
    const response = await patchSlideTextTool.execute("patch", {
      slide_index: 0,
      shape_id: "3",
      text: "Corrected translation",
      paragraph_start: 1,
      paragraph_end: 2,
      expected_text_hash: hashPlainText("Existing translation"),
    });
    const payload = JSON.parse(response.content[0].text);

    expect(payload).toMatchObject({
      success: true,
      shapeId: "3",
      verificationScope: { paragraph_start: 1, paragraph_end: 2 },
      verificationReadArgs: { paragraph_offset: 1, paragraph_limit: 1 },
    });
    expect(state.writes).toBe(1);
    expect(state.slideXml).toContain("Original source");
    expect(state.slideXml).toContain("Corrected translation");
    expect(state.slideXml).not.toContain("Existing translation");
  });
});
