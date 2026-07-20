import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import {
  applyPlainTextChange,
  applyPlainTextRangeChange,
  hashPlainText,
  inspectPlainTextSelection,
  MAX_OOXML_TEXT_WRITE_BYTES,
  MAX_PLAIN_TEXT_WRITE_BYTES,
  MAX_TEXT_RANGE_CHARACTERS,
  MAX_TEXT_RANGE_PARAGRAPHS,
  type PlainTextRange,
  utf8ByteLength,
} from "../pptx/text-xml";
import { findShapeById, sanitizeXmlAmpersands } from "../pptx/xml-utils";
import { unpackSlideZipResult } from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

interface RangeParams {
  paragraph_start?: number;
  paragraph_end?: number;
  char_start?: number;
  char_end?: number;
}

function parseRange(params: RangeParams): PlainTextRange | undefined {
  const hasRange =
    params.paragraph_start !== undefined ||
    params.paragraph_end !== undefined ||
    params.char_start !== undefined ||
    params.char_end !== undefined;
  if (!hasRange) return undefined;
  if (
    params.paragraph_start === undefined ||
    params.paragraph_end === undefined
  ) {
    throw new Error(
      "Range writes require both paragraph_start and paragraph_end",
    );
  }
  const hasCharStart = params.char_start !== undefined;
  const hasCharEnd = params.char_end !== undefined;
  if (hasCharStart !== hasCharEnd) {
    throw new Error("Character ranges require both char_start and char_end");
  }
  if (hasCharStart && hasCharEnd) {
    return {
      kind: "characters",
      paragraphStart: params.paragraph_start,
      paragraphEnd: params.paragraph_end,
      charStart: params.char_start as number,
      charEnd: params.char_end as number,
    };
  }
  return {
    kind: "paragraphs",
    paragraphStart: params.paragraph_start,
    paragraphEnd: params.paragraph_end,
  };
}

function assertInputBytes(value: string, maximum: number, label: string): void {
  const bytes = utf8ByteLength(value);
  if (bytes > maximum) {
    throw new Error(
      `${label} is ${bytes} UTF-8 bytes; the per-call limit is ${maximum}. Use read_slide_text pagination and bounded range writes.`,
    );
  }
}

export const editSlideTextTool = defineTool({
  name: "edit_slide_text",
  label: "Edit Slide Text",
  description:
    "Edit text in a shape. For ordinary text changes or translations, pass " +
    "text (and optionally mode=append) to preserve text-body, paragraph, and " +
    "first-run styling without reading OOXML. For paged text, pass the " +
    "paragraph/character scope and textHash returned by read_slide_text as " +
    "expected_text_hash. Range replace replaces the scope; range append inserts " +
    "after it. Mixed runs may be simplified; use raw OOXML for detailed edits.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    shape_id: Type.String({
      description:
        'Shape ID from list_slide_shapes or verify_slides output (e.g., "2", "20"). Stable and locale-independent.',
    }),
    code: Type.Optional(
      Type.String({
        description: `Raw OOXML <a:p> paragraph XML (maximum ${MAX_OOXML_TEXT_WRITE_BYTES} UTF-8 bytes; detailed formatting only)`,
        maxLength: MAX_OOXML_TEXT_WRITE_BYTES,
      }),
    ),
    text: Type.Optional(
      Type.String({
        description: `Plain text to replace or append (maximum ${MAX_PLAIN_TEXT_WRITE_BYTES} UTF-8 bytes). Preserves paragraph and first-run styling; use code for mixed-format runs.`,
        maxLength: MAX_PLAIN_TEXT_WRITE_BYTES,
      }),
    ),
    mode: Type.Optional(
      Type.Union([Type.Literal("replace"), Type.Literal("append")], {
        description: "Plain-text mode. Default replace.",
      }),
    ),
    paragraph_start: Type.Optional(
      Type.Integer({
        description:
          "0-based first paragraph from read_slide_text editScope (inclusive)",
        minimum: 0,
      }),
    ),
    paragraph_end: Type.Optional(
      Type.Integer({
        description: `0-based paragraph end from read_slide_text editScope (exclusive; maximum ${MAX_TEXT_RANGE_PARAGRAPHS} paragraphs per range)`,
        minimum: 1,
      }),
    ),
    char_start: Type.Optional(
      Type.Integer({
        description:
          "UTF-16 offset in paragraph_start from read_slide_text editScope (inclusive)",
        minimum: 0,
      }),
    ),
    char_end: Type.Optional(
      Type.Integer({
        description: `UTF-16 offset in paragraph_start from read_slide_text editScope (exclusive; maximum ${MAX_TEXT_RANGE_CHARACTERS} code units per range)`,
        minimum: 1,
      }),
    ),
    expected_text_hash: Type.Optional(
      Type.String({
        description:
          "textHash from read_slide_text. Required for range writes unless expected_text is supplied.",
        maxLength: 32,
      }),
    ),
    expected_text: Type.Optional(
      Type.String({
        description: `Exact original text in the scope (maximum ${MAX_PLAIN_TEXT_WRITE_BYTES} UTF-8 bytes). Alternative to expected_text_hash.`,
        maxLength: MAX_PLAIN_TEXT_WRITE_BYTES,
      }),
    ),
    explanation: Type.Optional(
      Type.String({
        description: "Brief description (max 50 chars)",
        maxLength: 50,
      }),
    ),
  }),
  execute: async (_toolCallId, params) => {
    try {
      if (params.code === undefined && params.text === undefined) {
        throw new Error("Provide either text or code");
      }
      if (params.code !== undefined && params.text !== undefined) {
        throw new Error("Provide text or code, not both");
      }
      const range = parseRange(params);
      if (params.code !== undefined && range) {
        throw new Error("Range writes support plain text only, not OOXML code");
      }
      if (
        params.code !== undefined &&
        (params.expected_text !== undefined ||
          params.expected_text_hash !== undefined)
      ) {
        throw new Error("Text guards support plain text writes only");
      }
      if (
        range &&
        params.expected_text === undefined &&
        params.expected_text_hash === undefined
      ) {
        throw new Error(
          "Range writes require expected_text_hash or expected_text from read_slide_text",
        );
      }
      if (params.text !== undefined) {
        assertInputBytes(params.text, MAX_PLAIN_TEXT_WRITE_BYTES, "Plain text");
      }
      if (params.code !== undefined) {
        assertInputBytes(params.code, MAX_OOXML_TEXT_WRITE_BYTES, "OOXML code");
      }
      if (params.expected_text !== undefined) {
        assertInputBytes(
          params.expected_text,
          MAX_PLAIN_TEXT_WRITE_BYTES,
          "expected_text",
        );
      }

      let rangeResult:
        | {
            beforeTextHash: string;
            afterTextHash: string;
            nextCursor: {
              paragraph_offset: number;
              char_offset?: number;
            } | null;
          }
        | undefined;
      let guardedWholeResult:
        | { beforeTextHash: string; afterTextHash: string }
        | undefined;
      const target = toSlideTargetReference(params);
      const zipValue = await safeRun(async (context) => {
        return withSlideZip(context, target, async ({ zip, markDirty }) => {
          const slideFile = zip.file("ppt/slides/slide1.xml");
          if (!slideFile) throw new Error("Slide XML not found in archive");

          const xml = await slideFile.async("string");
          const doc = new DOMParser().parseFromString(xml, "text/xml");

          const shape = findShapeById(doc, params.shape_id);
          if (!shape) {
            throw new Error(
              `Shape with id "${params.shape_id}" not found on the targeted slide. Use list_slide_shapes with the same slide_id to discover valid shape IDs.`,
            );
          }

          if (params.text !== undefined) {
            const selection = inspectPlainTextSelection(shape, range);
            if (
              params.expected_text !== undefined &&
              params.expected_text !== selection.text
            ) {
              throw new Error(
                `Text guard mismatch for requested scope (expected ${hashPlainText(params.expected_text)}, actual ${selection.textHash}). Re-read this scope before writing.`,
              );
            }
            if (
              params.expected_text_hash !== undefined &&
              params.expected_text_hash !== selection.textHash
            ) {
              throw new Error(
                `Text guard mismatch for requested scope (expected ${params.expected_text_hash}, actual ${selection.textHash}). Re-read this scope before writing.`,
              );
            }

            if (range) {
              rangeResult = applyPlainTextRangeChange(
                doc,
                shape,
                params.text,
                range,
                params.mode ?? "replace",
              );
            } else {
              applyPlainTextChange(
                doc,
                shape,
                params.text,
                params.mode ?? "replace",
              );
              if (
                params.expected_text !== undefined ||
                params.expected_text_hash !== undefined
              ) {
                guardedWholeResult = {
                  beforeTextHash: selection.textHash,
                  afterTextHash: inspectPlainTextSelection(shape).textHash,
                };
              }
            }
            zip.file(
              "ppt/slides/slide1.xml",
              new XMLSerializer().serializeToString(doc),
            );
            markDirty();
            return;
          }

          let txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];

          if (txBody) {
            const bodyPr = txBody.getElementsByTagNameNS(NS_A, "bodyPr")[0];
            const lstStyle = txBody.getElementsByTagNameNS(NS_A, "lstStyle")[0];

            while (txBody.firstChild) txBody.removeChild(txBody.firstChild);
            if (bodyPr) txBody.appendChild(bodyPr);
            if (lstStyle) txBody.appendChild(lstStyle);
          } else {
            txBody = doc.createElementNS(NS_P, "p:txBody");
            const bodyPr = doc.createElementNS(NS_A, "a:bodyPr");
            const lstStyle = doc.createElementNS(NS_A, "a:lstStyle");
            txBody.appendChild(bodyPr);
            txBody.appendChild(lstStyle);
            shape.appendChild(txBody);
          }

          const sanitizedXml = sanitizeXmlAmpersands(params.code ?? "");
          const wrapperXml = `<wrapper xmlns:a="${NS_A}" xmlns:r="${NS_R}">${sanitizedXml}</wrapper>`;
          const parsedDoc = new DOMParser().parseFromString(
            wrapperXml,
            "text/xml",
          );

          const parseError = parsedDoc.getElementsByTagName("parsererror")[0];
          if (parseError) {
            throw new Error(`Invalid XML: ${parseError.textContent}`);
          }

          const wrapper = parsedDoc.documentElement;
          let paragraphCount = 0;

          for (let i = 0; i < wrapper.childNodes.length; i++) {
            const child = wrapper.childNodes[i];
            if (child.nodeType === 1) {
              const el = child as Element;
              if (el.localName !== "p" || el.namespaceURI !== NS_A) {
                throw new Error(
                  `Invalid element <${el.nodeName}> — only <a:p> elements are allowed`,
                );
              }
              txBody.appendChild(doc.importNode(child, true));
              paragraphCount++;
            }
          }

          if (paragraphCount === 0) {
            throw new Error(
              "xml must contain at least one <a:p> paragraph element",
            );
          }

          zip.file(
            "ppt/slides/slide1.xml",
            new XMLSerializer().serializeToString(doc),
          );
          markDirty();
        });
      });
      const { metadata } = unpackSlideZipResult(zipValue, target);

      if (range && rangeResult) {
        return toolSuccess({
          success: true,
          ...metadata,
          ...(metadata.slideIndex === undefined &&
          params.slide_index !== undefined
            ? { slideIndex: params.slide_index }
            : {}),
          shapeId: params.shape_id,
          scope:
            range.kind === "characters"
              ? {
                  paragraph_start: range.paragraphStart,
                  paragraph_end: range.paragraphEnd,
                  char_start: range.charStart,
                  char_end: range.charEnd,
                }
              : {
                  paragraph_start: range.paragraphStart,
                  paragraph_end: range.paragraphEnd,
                },
          beforeTextHash: rangeResult.beforeTextHash,
          afterTextHash: rangeResult.afterTextHash,
          nextCursor: rangeResult.nextCursor,
          nextAction: rangeResult.nextCursor
            ? "Continue with read_slide_text using nextCursor."
            : "The edited scope reached the end of the shape.",
        });
      }
      if (guardedWholeResult) {
        return toolSuccess({
          success: true,
          ...metadata,
          ...guardedWholeResult,
        });
      }
      return toolSuccess({ success: true, ...metadata });
    } catch (error) {
      const normalized = toSlideMutationNotStartedError(
        error,
        "Failed to edit slide text",
      );
      return toolError(normalized.message, normalized);
    }
  },
  modifiedSlide: (params, result) =>
    result && typeof result.slideId === "string"
      ? result.slideId
      : (params.slide_id ?? params.slide_index),
});
