import { Window } from "happy-dom";
import { describe, expect, it } from "vitest";
import {
  applyPlainTextChange,
  applyPlainTextRangeChange,
  hashPlainText,
  inspectPlainTextSelection,
} from "../src/lib/pptx/text-xml";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";

function parseShape(texts = ["Hello"]) {
  const window = new Window();
  window.Element.prototype.getElementsByTagNameNS = function (
    namespace: string,
    localName: string,
  ) {
    return Array.from(this.getElementsByTagName("*")).filter(
      (element) =>
        element.namespaceURI === namespace && element.localName === localName,
    ) as unknown as HTMLCollectionOf<Element>;
  };
  const doc = window.document.implementation.createDocument(
    NS_P,
    "p:sld",
    null,
  ) as unknown as Document;
  const shape = doc.createElementNS(NS_P, "p:sp");
  const txBody = doc.createElementNS(NS_P, "p:txBody");
  txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"));
  txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
  for (const value of texts) {
    const paragraph = doc.createElementNS(NS_A, "a:p");
    const pPr = doc.createElementNS(NS_A, "a:pPr");
    pPr.setAttribute("lvl", "0");
    paragraph.appendChild(pPr);
    const run = doc.createElementNS(NS_A, "a:r");
    const rPr = doc.createElementNS(NS_A, "a:rPr");
    rPr.setAttribute("b", "1");
    run.appendChild(rPr);
    const text = doc.createElementNS(NS_A, "a:t");
    text.textContent = value;
    run.appendChild(text);
    paragraph.appendChild(run);
    paragraph.appendChild(doc.createElementNS(NS_A, "a:endParaRPr"));
    txBody.appendChild(paragraph);
  }
  shape.appendChild(txBody);
  doc.documentElement.appendChild(shape);
  return { doc, shape };
}

function shapeTexts(shape: Element): Array<string | null> {
  return Array.from(shape.getElementsByTagNameNS(NS_A, "t")).map(
    (node) => node.textContent,
  );
}

describe("applyPlainTextChange", () => {
  it("replaces text while preserving paragraph and first-run properties", () => {
    const { doc, shape } = parseShape();
    applyPlainTextChange(doc, shape, "Bonjour\nMonde");

    const paragraphs = shape.getElementsByTagNameNS(NS_A, "p");
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0].getElementsByTagNameNS(NS_A, "pPr")[0]).toBeTruthy();
    expect(
      paragraphs[0].getElementsByTagNameNS(NS_A, "rPr")[0].getAttribute("b"),
    ).toBe("1");
    expect(
      Array.from(shape.getElementsByTagNameNS(NS_A, "t")).map(
        (node) => node.textContent,
      ),
    ).toEqual(["Bonjour", "Monde"]);
  });

  it("appends translation paragraphs without replacing the source text", () => {
    const { doc, shape } = parseShape();
    applyPlainTextChange(doc, shape, "Bonjour", "append");

    expect(
      Array.from(shape.getElementsByTagNameNS(NS_A, "t")).map(
        (node) => node.textContent,
      ),
    ).toEqual(["Hello", "Bonjour"]);
  });
});

describe("applyPlainTextRangeChange", () => {
  it("replaces only the selected paragraph range and adjusts the next cursor", () => {
    const { doc, shape } = parseShape(["One", "Two", "Four"]);
    const result = applyPlainTextRangeChange(doc, shape, "Deux\nTrois", {
      kind: "paragraphs",
      paragraphStart: 1,
      paragraphEnd: 2,
    });

    expect(shapeTexts(shape)).toEqual(["One", "Deux", "Trois", "Four"]);
    expect(result.beforeTextHash).toBe(hashPlainText("Two"));
    expect(result.nextCursor).toEqual({ paragraph_offset: 3 });
  });

  it("inserts appended paragraphs after the selected range", () => {
    const { doc, shape } = parseShape(["Source", "Next"]);
    const result = applyPlainTextRangeChange(
      doc,
      shape,
      "Translation A\nTranslation B",
      { kind: "paragraphs", paragraphStart: 0, paragraphEnd: 1 },
      "append",
    );

    expect(shapeTexts(shape)).toEqual([
      "Source",
      "Translation A",
      "Translation B",
      "Next",
    ]);
    expect(result.nextCursor).toEqual({ paragraph_offset: 3 });
  });

  it("replaces a character range and moves the cursor by replacement length", () => {
    const { doc, shape } = parseShape(["Hello world"]);
    const range = {
      kind: "characters" as const,
      paragraphStart: 0,
      paragraphEnd: 1,
      charStart: 0,
      charEnd: 5,
    };
    expect(inspectPlainTextSelection(shape, range)).toMatchObject({
      text: "Hello",
      textHash: hashPlainText("Hello"),
    });

    const result = applyPlainTextRangeChange(doc, shape, "Bonjour", range);

    expect(shapeTexts(shape)).toEqual(["Bonjour world"]);
    expect(result.nextCursor).toEqual({
      paragraph_offset: 0,
      char_offset: 7,
    });
  });

  it("appends after a character range without deleting the source", () => {
    const { doc, shape } = parseShape(["Hello world"]);
    const result = applyPlainTextRangeChange(
      doc,
      shape,
      " / Bonjour",
      {
        kind: "characters",
        paragraphStart: 0,
        paragraphEnd: 1,
        charStart: 0,
        charEnd: 5,
      },
      "append",
    );

    expect(shapeTexts(shape)).toEqual(["Hello / Bonjour world"]);
    expect(result.nextCursor).toEqual({
      paragraph_offset: 0,
      char_offset: 15,
    });
  });

  it("rejects line breaks in a character-range write", () => {
    const { doc, shape } = parseShape();
    expect(() =>
      applyPlainTextRangeChange(doc, shape, "Bon\njour", {
        kind: "characters",
        paragraphStart: 0,
        paragraphEnd: 1,
        charStart: 0,
        charEnd: 5,
      }),
    ).toThrow("cannot contain line breaks");
  });

  it("rejects character boundaries that split a Unicode code point", () => {
    const { doc, shape } = parseShape(["A😀B"]);
    expect(() =>
      applyPlainTextRangeChange(doc, shape, "x", {
        kind: "characters",
        paragraphStart: 0,
        paragraphEnd: 1,
        charStart: 1,
        charEnd: 2,
      }),
    ).toThrow("must not split a Unicode surrogate pair");
  });
});
