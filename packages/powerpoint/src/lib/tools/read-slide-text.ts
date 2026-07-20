import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import {
  extractPlainParagraph,
  getTextParagraphs,
  hashPlainText,
  isUtf16CharacterBoundary,
  utf8ByteLength,
} from "../pptx/text-xml";
import { findShapeById } from "../pptx/xml-utils";
import {
  serializedJsonByteLength,
  TOOL_RESULT_MAX_BYTES,
} from "./result-budget";
import {
  attachSlideResultMetadata,
  unpackSlideZipResult,
} from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";

function truncateUtf8(
  text: string,
  maxBytes: number,
): {
  text: string;
  truncated: boolean;
  consumedChars: number;
} {
  const bytes = new TextEncoder().encode(text);
  if (bytes.byteLength <= maxBytes) {
    return { text, truncated: false, consumedChars: text.length };
  }
  let end = Math.max(0, maxBytes);
  while (end > 0) {
    try {
      const candidate = new TextDecoder("utf-8", { fatal: true }).decode(
        bytes.slice(0, end),
      );
      return {
        text: candidate,
        truncated: true,
        consumedChars: candidate.length,
      };
    } catch {
      end--;
    }
  }
  return { text: "", truncated: true, consumedChars: 0 };
}

export const readSlideTextTool = defineTool({
  name: "read_slide_text",
  label: "Read Slide Text",
  description:
    "Read text from a shape. Use format=plain for normal reading/translation " +
    "(shape text only, no fonts/colors/positions); use format=ooxml only when " +
    "detailed formatting, bullets, or styles are required.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    shape_id: Type.String({
      description:
        'Shape ID from list_slide_shapes or verify_slides output (e.g., "2", "20"). Stable and locale-independent.',
    }),
    format: Type.Optional(
      Type.Union([Type.Literal("plain"), Type.Literal("ooxml")], {
        description:
          'Output format. Default plain. Pass "ooxml" explicitly only when detailed formatting is required.',
        default: "plain",
      }),
    ),
    paragraph_offset: Type.Optional(
      Type.Number({ description: "0-based paragraph offset. Default 0." }),
    ),
    paragraph_limit: Type.Optional(
      Type.Number({
        description: "Maximum paragraphs to return. Default 100.",
      }),
    ),
    char_offset: Type.Optional(
      Type.Number({
        description:
          "0-based character offset within the first plain-text paragraph page.",
      }),
    ),
    max_bytes: Type.Optional(
      Type.Number({
        description:
          "Maximum text/XML UTF-8 bytes before JSON wrapping. Default and max 8000.",
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
      const target = toSlideTargetReference(params);
      const zipValue = await safeRun(async (context) =>
        withSlideZip(context, target, async ({ zip }) => {
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

          const txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];
          if (!txBody) {
            return "(empty — shape has no text body)";
          }

          const paragraphs = getTextParagraphs(txBody);
          if (paragraphs.length === 0) {
            return "(empty — shape has a text body but no paragraph content)";
          }
          const offset = Math.max(0, Math.floor(params.paragraph_offset ?? 0));
          const limit = Math.min(
            200,
            Math.max(1, Math.floor(params.paragraph_limit ?? 100)),
          );
          const selected = paragraphs.slice(offset, offset + limit);
          const maxBytes = Math.min(
            8_000,
            Math.max(1_000, Math.floor(params.max_bytes ?? 8_000)),
          );

          if ((params.format ?? "plain") === "plain") {
            const charOffset = Math.max(0, Math.floor(params.char_offset ?? 0));
            type PlainTextItem = {
              index: number;
              text: string;
              textHash: string;
              editScope:
                | {
                    paragraph_start: number;
                    paragraph_end: number;
                  }
                | {
                    paragraph_start: number;
                    paragraph_end: number;
                    char_start: number;
                    char_end: number;
                  };
              textTruncated?: boolean;
            };
            const plainItems: PlainTextItem[] = [];
            let usedBytes = 0;
            let nextOffset: number | null = null;
            let nextCharOffset: number | null = null;
            let budgetLimited = false;

            for (let i = 0; i < selected.length; i++) {
              const fullText = extractPlainParagraph(selected[i]);
              const startChar = i === 0 ? charOffset : 0;
              if (startChar > fullText.length) {
                throw new Error(
                  `char_offset ${startChar} exceeds paragraph ${offset + i} length ${fullText.length}`,
                );
              }
              if (!isUtf16CharacterBoundary(fullText, startChar)) {
                throw new Error(
                  `char_offset ${startChar} splits a Unicode surrogate pair in paragraph ${offset + i}`,
                );
              }
              if (i === 0 && startChar === fullText.length && startChar > 0) {
                nextOffset = offset + 1 < paragraphs.length ? offset + 1 : null;
                continue;
              }
              const text = fullText.slice(startChar);
              const remainingBytes = Math.max(0, maxBytes - usedBytes);
              if (remainingBytes === 0) {
                nextOffset = offset + i;
                nextCharOffset = startChar > 0 ? startChar : null;
                break;
              }
              const capped = truncateUtf8(text, remainingBytes);
              const paragraphIndex = offset + i;
              if (capped.truncated && capped.consumedChars === 0) {
                nextOffset = paragraphIndex;
                nextCharOffset = startChar > 0 ? startChar : null;
                break;
              }
              const isCharacterRange = startChar > 0 || capped.truncated;
              plainItems.push({
                index: paragraphIndex,
                text: capped.text,
                textHash: hashPlainText(capped.text),
                editScope: isCharacterRange
                  ? {
                      paragraph_start: paragraphIndex,
                      paragraph_end: paragraphIndex + 1,
                      char_start: startChar,
                      char_end: startChar + capped.consumedChars,
                    }
                  : {
                      paragraph_start: paragraphIndex,
                      paragraph_end: paragraphIndex + 1,
                    },
                ...(capped.truncated ? { textTruncated: true } : {}),
              });
              usedBytes += utf8ByteLength(capped.text);
              if (capped.truncated) {
                nextOffset = offset + i;
                nextCharOffset = startChar + capped.consumedChars;
                break;
              }
              const followingParagraph = offset + i + 1;
              nextOffset =
                followingParagraph < paragraphs.length
                  ? followingParagraph
                  : null;
            }

            const buildPlainResult = () => {
              const isCompleteParagraphPage =
                plainItems.length > 0 &&
                plainItems.every((item) => !("char_start" in item.editScope));
              const pageEditScope = isCompleteParagraphPage
                ? {
                    paragraph_start: plainItems[0].index,
                    paragraph_end: plainItems[plainItems.length - 1].index + 1,
                  }
                : null;
              const isCompleteShapeRead =
                offset === 0 &&
                charOffset === 0 &&
                nextOffset === null &&
                plainItems.length === paragraphs.length;
              return {
                schemaVersion: 2,
                format: "plain",
                slideIndex: params.slide_index,
                shapeId: params.shape_id,
                paragraphs: plainItems,
                paragraphCount: paragraphs.length,
                ...(isCompleteShapeRead
                  ? {
                      shapeTextHash: hashPlainText(
                        plainItems.map((item) => item.text).join("\n"),
                      ),
                    }
                  : {}),
                page: {
                  offset,
                  limit,
                  returned: plainItems.length,
                  hasMore: nextOffset !== null,
                  nextOffset,
                  nextCharOffset,
                  nextCursor:
                    nextOffset === null
                      ? null
                      : {
                          paragraph_offset: nextOffset,
                          ...(nextCharOffset === null
                            ? {}
                            : { char_offset: nextCharOffset }),
                        },
                  budgetLimited,
                  ...(pageEditScope
                    ? {
                        editScope: pageEditScope,
                        textHash: hashPlainText(
                          plainItems.map((item) => item.text).join("\n"),
                        ),
                      }
                    : {}),
                },
                truncated: nextCharOffset !== null,
                omittedFields: [
                  "font",
                  "fontSize",
                  "color",
                  "position",
                  "geometry",
                  "rawOoxml",
                ],
              };
            };

            let plainResult = buildPlainResult();
            const isWithinToolBudget = () =>
              serializedJsonByteLength({
                success: true,
                result: buildPlainResult(),
              }) <=
              TOOL_RESULT_MAX_BYTES - 2_048;

            while (!isWithinToolBudget() && plainItems.length > 0) {
              budgetLimited = true;
              const itemIndex = plainItems.length - 1;
              const originalItem = plainItems[itemIndex];
              if (
                plainItems.length > 1 &&
                !("char_start" in originalItem.editScope)
              ) {
                plainItems.pop();
                nextOffset = originalItem.index;
                nextCharOffset = null;
                plainResult = buildPlainResult();
                continue;
              }
              const sourceCharStart =
                "char_start" in originalItem.editScope
                  ? originalItem.editScope.char_start
                  : 0;
              const boundaries = [0];
              for (let index = 1; index <= originalItem.text.length; index++) {
                if (isUtf16CharacterBoundary(originalItem.text, index)) {
                  boundaries.push(index);
                }
              }

              const makeTruncatedItem = (length: number): PlainTextItem => {
                const text = originalItem.text.slice(0, length);
                return {
                  index: originalItem.index,
                  text,
                  textHash: hashPlainText(text),
                  editScope: {
                    paragraph_start: originalItem.index,
                    paragraph_end: originalItem.index + 1,
                    char_start: sourceCharStart,
                    char_end: sourceCharStart + length,
                  },
                  textTruncated: true,
                };
              };

              let low = 1;
              let high = boundaries.length - 1;
              let bestLength = 0;
              while (low <= high) {
                const middle = Math.floor((low + high) / 2);
                const candidateLength = boundaries[middle];
                plainItems[itemIndex] = makeTruncatedItem(candidateLength);
                nextOffset = originalItem.index;
                nextCharOffset = sourceCharStart + candidateLength;
                if (isWithinToolBudget()) {
                  bestLength = candidateLength;
                  low = middle + 1;
                } else {
                  high = middle - 1;
                }
              }

              if (bestLength > 0) {
                plainItems[itemIndex] = makeTruncatedItem(bestLength);
                nextOffset = originalItem.index;
                nextCharOffset = sourceCharStart + bestLength;
              } else {
                plainItems.pop();
                nextOffset = originalItem.index;
                nextCharOffset = sourceCharStart > 0 ? sourceCharStart : null;
              }
              plainResult = buildPlainResult();
            }

            return plainResult;
          }

          const serializer = new XMLSerializer();
          const raw = selected
            .map((paragraph) => serializer.serializeToString(paragraph))
            .join("\n");
          const rawBytes = new TextEncoder().encode(raw).byteLength;
          const nextOffset =
            offset + selected.length < paragraphs.length
              ? offset + selected.length
              : null;
          if (rawBytes <= maxBytes) {
            // Preserve the legacy bare-string result only when the shape was
            // read in full. Paged reads must expose their continuation cursor.
            if (offset === 0 && nextOffset === null) return raw;
            return {
              schemaVersion: 2,
              format: "ooxml",
              slideIndex: params.slide_index,
              shapeId: params.shape_id,
              xml: raw,
              paragraphCount: paragraphs.length,
              page: {
                offset,
                limit,
                returned: selected.length,
                hasMore: nextOffset !== null,
                nextOffset,
              },
              truncated: false,
            };
          }
          return {
            schemaVersion: 2,
            format: "ooxml",
            slideIndex: params.slide_index,
            shapeId: params.shape_id,
            xml: null,
            paragraphCount: paragraphs.length,
            page: {
              offset,
              limit,
              returned: 0,
              hasMore: true,
              nextOffset: offset,
            },
            truncated: true,
            omittedBytes: rawBytes,
            suggestedParagraphLimit: Math.max(
              1,
              Math.floor(selected.length / 2),
            ),
            nextAction:
              selected.length <= 1
                ? "This paragraph exceeds the OOXML response budget. Use format=plain unless detailed run formatting is essential; OOXML is never returned partially."
                : "Reduce paragraph_limit or increase paragraph_offset. OOXML is never returned partially because truncated XML would be invalid.",
          };
        }),
      );
      const { result, metadata } = unpackSlideZipResult(zipValue, target);
      return toolSuccess({
        success: true,
        result: attachSlideResultMetadata(result, metadata),
        ...metadata,
      });
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "Failed to read slide text";
      return toolError(message);
    }
  },
});
