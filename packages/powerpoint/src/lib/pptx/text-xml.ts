const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";

export const MAX_PLAIN_TEXT_WRITE_BYTES = 16_000;
export const MAX_OOXML_TEXT_WRITE_BYTES = 64_000;
export const MAX_BATCH_PLAIN_TEXT_WRITE_BYTES = 64_000;
export const MAX_TEXT_RANGE_PARAGRAPHS = 200;
export const MAX_TEXT_RANGE_CHARACTERS = 16_000;

export interface ParagraphTextRange {
  kind: "paragraphs";
  paragraphStart: number;
  paragraphEnd: number;
}

export interface CharacterTextRange {
  kind: "characters";
  paragraphStart: number;
  paragraphEnd: number;
  charStart: number;
  charEnd: number;
}

export type PlainTextRange = ParagraphTextRange | CharacterTextRange;

export interface PlainTextSelection {
  text: string;
  textHash: string;
  paragraphCount: number;
}

export interface PlainTextRangeChangeResult {
  beforeTextHash: string;
  afterTextHash: string;
  nextCursor: {
    paragraph_offset: number;
    char_offset?: number;
  } | null;
}

export function utf8ByteLength(value: string): number {
  return new TextEncoder().encode(value).byteLength;
}

/** Small deterministic checksum used as an optimistic-concurrency guard. */
export function hashPlainText(value: string): string {
  let hash = 2166136261;
  for (const byte of new TextEncoder().encode(value)) {
    hash ^= byte;
    hash = Math.imul(hash, 16777619);
  }
  return `fnv1a32:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

export function getTextParagraphs(txBody: Element): Element[] {
  return Array.from(txBody.childNodes).filter(
    (node): node is Element =>
      node.nodeType === 1 &&
      (node as Element).localName === "p" &&
      (node as Element).namespaceURI === NS_A,
  );
}

export function extractPlainParagraph(paragraph: Element): string {
  let text = "";
  const visit = (node: Node) => {
    if (node.nodeType !== 1) return;
    const element = node as Element;
    if (element.localName === "t") {
      text += element.textContent ?? "";
      return;
    }
    if (element.localName === "br") {
      text += "\n";
      return;
    }
    if (element.localName === "tab") {
      text += "\t";
      return;
    }
    for (const child of Array.from(element.childNodes)) visit(child);
  };
  visit(paragraph);
  return text;
}

function ensureTextBody(doc: Document, shape: Element): Element {
  let txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];
  if (txBody) return txBody;

  txBody = doc.createElementNS(NS_P, "p:txBody");
  txBody.appendChild(doc.createElementNS(NS_A, "a:bodyPr"));
  txBody.appendChild(doc.createElementNS(NS_A, "a:lstStyle"));
  shape.appendChild(txBody);
  return txBody;
}

function createParagraph(
  doc: Document,
  source: Element | undefined,
  text: string,
): Element {
  const paragraph = source
    ? (source.cloneNode(true) as Element)
    : doc.createElementNS(NS_A, "a:p");
  const pPr = Array.from(paragraph.childNodes).find(
    (node): node is Element =>
      node.nodeType === 1 && (node as Element).localName === "pPr",
  );
  const endParaRPr = Array.from(paragraph.childNodes).find(
    (node): node is Element =>
      node.nodeType === 1 && (node as Element).localName === "endParaRPr",
  );
  const firstRun = Array.from(paragraph.getElementsByTagNameNS(NS_A, "r"))[0];
  const runProperties = firstRun?.getElementsByTagNameNS(NS_A, "rPr")[0];

  while (paragraph.firstChild) paragraph.removeChild(paragraph.firstChild);
  if (pPr) paragraph.appendChild(pPr);

  const run = doc.createElementNS(NS_A, "a:r");
  if (runProperties) run.appendChild(runProperties.cloneNode(true));
  const textNode = doc.createElementNS(NS_A, "a:t");
  textNode.textContent = text;
  if (/^\s|\s$|\s{2,}/.test(text)) {
    textNode.setAttributeNS(
      "http://www.w3.org/XML/1998/namespace",
      "xml:space",
      "preserve",
    );
  }
  run.appendChild(textNode);
  paragraph.appendChild(run);
  if (endParaRPr) paragraph.appendChild(endParaRPr);
  return paragraph;
}

/** Replace or append plain text while keeping paragraph/first-run styling. */
export function applyPlainTextChange(
  doc: Document,
  shape: Element,
  text: string,
  mode: "replace" | "append" = "replace",
): void {
  const txBody = ensureTextBody(doc, shape);
  const paragraphs = getTextParagraphs(txBody);
  const lines = text.replace(/\r\n?/g, "\n").split("\n");

  if (mode === "append") {
    const source = paragraphs[paragraphs.length - 1];
    for (const line of lines) {
      txBody.appendChild(createParagraph(doc, source, line));
    }
    return;
  }

  const bodyPr = txBody.getElementsByTagNameNS(NS_A, "bodyPr")[0];
  const lstStyle = txBody.getElementsByTagNameNS(NS_A, "lstStyle")[0];
  while (txBody.firstChild) txBody.removeChild(txBody.firstChild);
  if (bodyPr) txBody.appendChild(bodyPr);
  if (lstStyle) txBody.appendChild(lstStyle);

  for (let i = 0; i < lines.length; i++) {
    txBody.appendChild(
      createParagraph(
        doc,
        paragraphs[i] ?? paragraphs[paragraphs.length - 1],
        lines[i],
      ),
    );
  }
}

function assertInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${label} must be a non-negative integer`);
  }
}

export function isUtf16CharacterBoundary(
  text: string,
  offset: number,
): boolean {
  if (offset <= 0 || offset >= text.length) return true;
  const previous = text.charCodeAt(offset - 1);
  const current = text.charCodeAt(offset);
  return !(
    previous >= 0xd800 &&
    previous <= 0xdbff &&
    current >= 0xdc00 &&
    current <= 0xdfff
  );
}

function inspectRange(
  paragraphs: Element[],
  range: PlainTextRange,
): PlainTextSelection {
  assertInteger(range.paragraphStart, "paragraph_start");
  assertInteger(range.paragraphEnd, "paragraph_end");
  if (range.paragraphEnd <= range.paragraphStart) {
    throw new Error("paragraph_end must be greater than paragraph_start");
  }
  if (range.paragraphEnd > paragraphs.length) {
    throw new Error(
      `Paragraph range [${range.paragraphStart}, ${range.paragraphEnd}) exceeds paragraph count ${paragraphs.length}`,
    );
  }
  if (range.paragraphEnd - range.paragraphStart > MAX_TEXT_RANGE_PARAGRAPHS) {
    throw new Error(
      `A range write can touch at most ${MAX_TEXT_RANGE_PARAGRAPHS} paragraphs`,
    );
  }

  if (range.kind === "paragraphs") {
    const text = paragraphs
      .slice(range.paragraphStart, range.paragraphEnd)
      .map(extractPlainParagraph)
      .join("\n");
    return {
      text,
      textHash: hashPlainText(text),
      paragraphCount: paragraphs.length,
    };
  }

  if (range.paragraphEnd !== range.paragraphStart + 1) {
    throw new Error("A character range must stay within one paragraph");
  }
  assertInteger(range.charStart, "char_start");
  assertInteger(range.charEnd, "char_end");
  if (range.charEnd <= range.charStart) {
    throw new Error("char_end must be greater than char_start");
  }
  if (range.charEnd - range.charStart > MAX_TEXT_RANGE_CHARACTERS) {
    throw new Error(
      `A character-range write can touch at most ${MAX_TEXT_RANGE_CHARACTERS} UTF-16 code units`,
    );
  }

  const paragraphText = extractPlainParagraph(paragraphs[range.paragraphStart]);
  if (range.charEnd > paragraphText.length) {
    throw new Error(
      `Character range [${range.charStart}, ${range.charEnd}) exceeds paragraph ${range.paragraphStart} length ${paragraphText.length}`,
    );
  }
  if (
    !isUtf16CharacterBoundary(paragraphText, range.charStart) ||
    !isUtf16CharacterBoundary(paragraphText, range.charEnd)
  ) {
    throw new Error(
      "Character range boundaries must not split a Unicode surrogate pair",
    );
  }
  const text = paragraphText.slice(range.charStart, range.charEnd);
  return {
    text,
    textHash: hashPlainText(text),
    paragraphCount: paragraphs.length,
  };
}

export function inspectPlainTextSelection(
  shape: Element,
  range?: PlainTextRange,
): PlainTextSelection {
  const txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];
  const paragraphs = txBody ? getTextParagraphs(txBody) : [];
  if (range) return inspectRange(paragraphs, range);

  const text = paragraphs.map(extractPlainParagraph).join("\n");
  return {
    text,
    textHash: hashPlainText(text),
    paragraphCount: paragraphs.length,
  };
}

/**
 * Replace a bounded range, or insert text immediately after it. Character
 * writes intentionally simplify only the touched paragraph to one styled run.
 */
export function applyPlainTextRangeChange(
  doc: Document,
  shape: Element,
  text: string,
  range: PlainTextRange,
  mode: "replace" | "append" = "replace",
): PlainTextRangeChangeResult {
  const txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];
  if (!txBody) throw new Error("Shape has no text body");
  const paragraphs = getTextParagraphs(txBody);
  const selection = inspectRange(paragraphs, range);
  const hasFollowingParagraph = range.paragraphEnd < paragraphs.length;

  if (range.kind === "characters") {
    if (/\r|\n/.test(text)) {
      throw new Error(
        "Character-range text cannot contain line breaks; use a paragraph range instead",
      );
    }
    const source = paragraphs[range.paragraphStart];
    const original = extractPlainParagraph(source);
    const before = original.slice(0, range.charStart);
    const selected = original.slice(range.charStart, range.charEnd);
    const after = original.slice(range.charEnd);
    const replacement =
      mode === "append"
        ? `${before}${selected}${text}${after}`
        : `${before}${text}${after}`;
    txBody.replaceChild(createParagraph(doc, source, replacement), source);

    const nextCursor =
      range.charEnd < original.length
        ? {
            paragraph_offset: range.paragraphStart,
            char_offset:
              mode === "append"
                ? range.charEnd + text.length
                : range.charStart + text.length,
          }
        : hasFollowingParagraph
          ? { paragraph_offset: range.paragraphStart + 1 }
          : null;
    const changedText = mode === "append" ? `${selected}${text}` : text;
    return {
      beforeTextHash: selection.textHash,
      afterTextHash: hashPlainText(changedText),
      nextCursor,
    };
  }

  const normalized = text.replace(/\r\n?/g, "\n");
  const lines = normalized.split("\n");
  const selected = paragraphs.slice(range.paragraphStart, range.paragraphEnd);
  const insertionReference = paragraphs[range.paragraphEnd] ?? null;
  const styleSources =
    mode === "append" ? [selected[selected.length - 1]] : selected;
  const replacements = lines.map((line, index) =>
    createParagraph(
      doc,
      styleSources[index] ?? styleSources[styleSources.length - 1],
      line,
    ),
  );

  if (mode === "replace") {
    for (const paragraph of selected) txBody.removeChild(paragraph);
  }
  for (const replacement of replacements) {
    txBody.insertBefore(replacement, insertionReference);
  }

  const nextParagraph =
    mode === "replace"
      ? range.paragraphStart + replacements.length
      : range.paragraphEnd + replacements.length;
  return {
    beforeTextHash: selection.textHash,
    afterTextHash: hashPlainText(
      mode === "append" ? `${selection.text}\n${normalized}` : normalized,
    ),
    nextCursor: hasFollowingParagraph
      ? { paragraph_offset: nextParagraph }
      : null,
  };
}
