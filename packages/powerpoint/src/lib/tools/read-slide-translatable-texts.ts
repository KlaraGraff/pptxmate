import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import {
  extractPlainParagraph,
  hashPlainText,
  isUtf16CharacterBoundary,
  utf8ByteLength,
} from "../pptx/text-xml";
import {
  attachSlideResultMetadata,
  unpackSlideZipResult,
} from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SLIDE_RELS_PATH = "ppt/slides/_rels/slide1.xml.rels";

type TranslatableContentKind = "shape" | "groupShape" | "tableCell" | "chart";

interface TranslatableContentSource {
  kind: TranslatableContentKind;
  location: Record<string, unknown>;
  text: string;
}

interface UnsupportedContentSource {
  shapeId: string;
  groupShapeIds: string[];
  type: "graphicFrame" | "chart";
  reason: string;
}

function elementChildren(element: Element): Element[] {
  return Array.from(element.childNodes).filter(
    (node): node is Element => node.nodeType === 1,
  );
}

function getShapeIdentity(shape: Element): { id: string; name: string } {
  const cNvPr = shape.getElementsByTagNameNS(NS_P, "cNvPr")[0];
  return {
    id: cNvPr?.getAttribute("id") ?? "",
    name: cNvPr?.getAttribute("name") ?? "",
  };
}

function getTextFromBody(textBody: Element | undefined): string | null {
  if (!textBody) return null;
  const paragraphs = elementChildren(textBody).filter(
    (element) => element.namespaceURI === NS_A && element.localName === "p",
  );
  return paragraphs.map(extractPlainParagraph).join("\n");
}

function getShapeText(shape: Element): string | null {
  return getTextFromBody(shape.getElementsByTagNameNS(NS_P, "txBody")[0]);
}

function getTableCellText(cell: Element): string | null {
  return getTextFromBody(cell.getElementsByTagNameNS(NS_A, "txBody")[0]);
}

function resolvePartPath(sourcePart: string, target: string): string | null {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return null;
  const segments = target.startsWith("/")
    ? target.slice(1).split("/")
    : [...sourcePart.split("/").slice(0, -1), ...target.split("/")];
  const normalized: string[] = [];
  for (const segment of segments) {
    if (!segment || segment === ".") continue;
    if (segment === "..") {
      normalized.pop();
    } else {
      normalized.push(segment);
    }
  }
  return normalized.join("/");
}

async function chartRelationshipTargets(
  zip: import("jszip"),
): Promise<Map<string, string>> {
  const relsFile = zip.file(SLIDE_RELS_PATH);
  if (!relsFile) return new Map();
  const relsXml = await relsFile.async("string");
  const relsDoc = new DOMParser().parseFromString(relsXml, "text/xml");
  const targets = new Map<string, string>();
  for (const relationship of Array.from(
    relsDoc.getElementsByTagName("Relationship"),
  )) {
    const id = relationship.getAttribute("Id");
    const target = relationship.getAttribute("Target");
    if (!id || !target) continue;
    const path = resolvePartPath("ppt/slides/slide1.xml", target);
    if (path) targets.set(id, path);
  }
  return targets;
}

function extractChartText(chartDoc: Document): string {
  const drawingText = Array.from(
    chartDoc.getElementsByTagNameNS(NS_A, "t"),
  ).map((element) => element.textContent ?? "");
  const cachedValues = Array.from(
    chartDoc.getElementsByTagNameNS(NS_C, "v"),
  ).map((element) => element.textContent ?? "");
  return [...drawingText, ...cachedValues].filter(Boolean).join("\n");
}

async function collectGraphicFrameContent(
  frame: Element,
  groupShapeIds: string[],
  zip: import("jszip"),
  chartTargets: Map<string, string>,
  includeEmpty: boolean,
  sources: TranslatableContentSource[],
  unsupported: UnsupportedContentSource[],
): Promise<void> {
  const identity = getShapeIdentity(frame);
  const graphicData = frame.getElementsByTagNameNS(NS_A, "graphicData")[0];
  const uri = graphicData?.getAttribute("uri") ?? "";

  if (uri.includes("/table")) {
    const table = frame.getElementsByTagNameNS(NS_A, "tbl")[0];
    if (!table) {
      unsupported.push({
        shapeId: identity.id,
        groupShapeIds,
        type: "graphicFrame",
        reason: "The table graphic frame has no readable table XML.",
      });
      return;
    }
    const rows = Array.from(table.getElementsByTagNameNS(NS_A, "tr"));
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
      const cells = elementChildren(rows[rowIndex]).filter(
        (cell) => cell.namespaceURI === NS_A && cell.localName === "tc",
      );
      for (let columnIndex = 0; columnIndex < cells.length; columnIndex++) {
        const text = getTableCellText(cells[columnIndex]);
        if (text === null || (!includeEmpty && text.length === 0)) continue;
        sources.push({
          kind: "tableCell",
          location: {
            shapeId: identity.id,
            shapeName: identity.name,
            groupShapeIds,
            rowIndex,
            columnIndex,
          },
          text,
        });
      }
    }
    return;
  }

  if (uri.includes("/chart")) {
    const chartReferences = Array.from(
      frame.getElementsByTagNameNS(NS_C, "chart"),
    );
    if (chartReferences.length === 0) {
      unsupported.push({
        shapeId: identity.id,
        groupShapeIds,
        type: "chart",
        reason: "The chart graphic frame has no chart relationship.",
      });
      return;
    }
    for (
      let chartIndex = 0;
      chartIndex < chartReferences.length;
      chartIndex++
    ) {
      const chartReference = chartReferences[chartIndex];
      const relationshipId =
        chartReference.getAttributeNS(NS_R, "id") ??
        chartReference.getAttribute("r:id");
      const chartPath = relationshipId
        ? chartTargets.get(relationshipId)
        : undefined;
      const chartFile = chartPath ? zip.file(chartPath) : null;
      if (!chartFile) {
        unsupported.push({
          shapeId: identity.id,
          groupShapeIds,
          type: "chart",
          reason: relationshipId
            ? `Chart relationship "${relationshipId}" has no readable chart XML.`
            : "Chart relationship ID is missing.",
        });
        continue;
      }
      const chartXml = await chartFile.async("string");
      const chartDoc = new DOMParser().parseFromString(chartXml, "text/xml");
      const text = extractChartText(chartDoc);
      if (!includeEmpty && text.length === 0) continue;
      sources.push({
        kind: "chart",
        location: {
          shapeId: identity.id,
          shapeName: identity.name,
          groupShapeIds,
          chartIndex,
          chartPath,
        },
        text,
      });
    }
    return;
  }

  unsupported.push({
    shapeId: identity.id,
    groupShapeIds,
    type: "graphicFrame",
    reason: `Unsupported graphic-frame URI "${uri || "unknown"}".`,
  });
}

async function collectTranslatableContent(
  doc: Document,
  zip: import("jszip"),
  includeEmpty: boolean,
): Promise<{
  sources: TranslatableContentSource[];
  unsupported: UnsupportedContentSource[];
}> {
  const sources: TranslatableContentSource[] = [];
  const unsupported: UnsupportedContentSource[] = [];
  const chartTargets = await chartRelationshipTargets(zip);
  const shapeTree = doc.getElementsByTagNameNS(NS_P, "spTree")[0];
  if (!shapeTree) return { sources, unsupported };

  const visit = async (container: Element, groupShapeIds: string[]) => {
    for (const child of elementChildren(container)) {
      if (child.namespaceURI !== NS_P) continue;
      if (child.localName === "sp") {
        const identity = getShapeIdentity(child);
        const text = getShapeText(child);
        if (text === null || (!includeEmpty && text.length === 0)) continue;
        sources.push({
          kind: groupShapeIds.length > 0 ? "groupShape" : "shape",
          location: {
            shapeId: identity.id,
            shapeName: identity.name,
            ...(groupShapeIds.length > 0 ? { groupShapeIds } : {}),
          },
          text,
        });
      } else if (child.localName === "grpSp") {
        const identity = getShapeIdentity(child);
        await visit(child, [...groupShapeIds, identity.id]);
      } else if (child.localName === "graphicFrame") {
        await collectGraphicFrameContent(
          child,
          groupShapeIds,
          zip,
          chartTargets,
          includeEmpty,
          sources,
          unsupported,
        );
      }
    }
  };

  await visit(shapeTree, []);
  return { sources, unsupported };
}

function truncateUtf8(
  text: string,
  maxBytes: number,
): { text: string; truncated: boolean; consumedChars: number } {
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

export const readSlideTranslatableTextsTool = defineTool({
  name: "read_slide_translatable_texts",
  label: "Read Slide Translatable Texts",
  description:
    "Inventory all readable translatable text on one slide for translation audit. " +
    "Covers normal shapes, grouped shapes, table cells, and chart XML. Results " +
    "are paginated by a required continuation cursor; do not claim a complete " +
    "audit while hasMore is true or unsupported containers are reported.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    offset: Type.Optional(
      Type.Number({ description: "0-based content-source offset. Default 0." }),
    ),
    char_offset: Type.Optional(
      Type.Number({
        description:
          "UTF-16 offset within the source at offset. Use page.nextCursor.char_offset.",
      }),
    ),
    limit: Type.Optional(
      Type.Number({
        description: "Maximum source chunks. Default 50, max 100.",
      }),
    ),
    max_bytes: Type.Optional(
      Type.Number({
        description:
          "Maximum UTF-8 text budget before JSON wrapping. Default and max 8000.",
      }),
    ),
    include_empty: Type.Optional(
      Type.Boolean({
        description: "Include empty readable text containers. Default false.",
      }),
    ),
    explanation: Type.Optional(
      Type.String({
        description: "Brief description (max 80 chars)",
        maxLength: 80,
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
          const slideXml = await slideFile.async("string");
          const doc = new DOMParser().parseFromString(slideXml, "text/xml");
          const { sources, unsupported } = await collectTranslatableContent(
            doc,
            zip,
            params.include_empty === true,
          );

          const offset = Math.max(0, Math.floor(params.offset ?? 0));
          if (offset > sources.length) {
            throw new Error(
              `offset ${offset} exceeds translatable content count ${sources.length}`,
            );
          }
          const initialCharOffset = Math.max(
            0,
            Math.floor(params.char_offset ?? 0),
          );
          const limit = Math.min(
            100,
            Math.max(1, Math.floor(params.limit ?? 50)),
          );
          const maxBytes = Math.min(
            8_000,
            Math.max(1_000, Math.floor(params.max_bytes ?? 8_000)),
          );
          const items: Array<{
            sourceIndex: number;
            kind: TranslatableContentKind;
            location: Record<string, unknown>;
            text: string;
            textHash: string;
            span: { charStart: number; charEnd: number };
            textTruncated?: boolean;
          }> = [];
          let sourceIndex = offset;
          let charOffset = initialCharOffset;
          let usedBytes = 0;
          let nextCursor: { offset: number; char_offset?: number } | null =
            null;

          while (sourceIndex < sources.length && items.length < limit) {
            const source = sources[sourceIndex];
            if (charOffset > source.text.length) {
              throw new Error(
                `char_offset ${charOffset} exceeds source ${sourceIndex} length ${source.text.length}`,
              );
            }
            if (!isUtf16CharacterBoundary(source.text, charOffset)) {
              throw new Error(
                `char_offset ${charOffset} splits a Unicode surrogate pair in source ${sourceIndex}`,
              );
            }
            const fixedBytes =
              utf8ByteLength(
                JSON.stringify({
                  sourceIndex,
                  kind: source.kind,
                  location: source.location,
                  span: { charStart: charOffset, charEnd: source.text.length },
                }),
              ) + 80;
            const available = maxBytes - usedBytes - fixedBytes;
            if (available <= 0) {
              nextCursor = {
                offset: sourceIndex,
                ...(charOffset > 0 ? { char_offset: charOffset } : {}),
              };
              break;
            }
            const remainder = source.text.slice(charOffset);
            const capped = truncateUtf8(remainder, available);
            if (remainder.length > 0 && capped.consumedChars === 0) {
              nextCursor = {
                offset: sourceIndex,
                ...(charOffset > 0 ? { char_offset: charOffset } : {}),
              };
              break;
            }
            const charEnd = charOffset + capped.consumedChars;
            items.push({
              sourceIndex,
              kind: source.kind,
              location: source.location,
              text: capped.text,
              textHash: hashPlainText(capped.text),
              span: { charStart: charOffset, charEnd },
              ...(capped.truncated ? { textTruncated: true } : {}),
            });
            usedBytes += fixedBytes + utf8ByteLength(capped.text);
            if (capped.truncated) {
              nextCursor = { offset: sourceIndex, char_offset: charEnd };
              break;
            }
            sourceIndex++;
            charOffset = 0;
          }

          if (!nextCursor && sourceIndex < sources.length) {
            nextCursor = { offset: sourceIndex };
          }
          const sourceCounts = sources.reduce<Record<string, number>>(
            (counts, source) => {
              counts[source.kind] = (counts[source.kind] ?? 0) + 1;
              return counts;
            },
            {},
          );
          return {
            schemaVersion: 1,
            items,
            page: {
              offset,
              charOffset: initialCharOffset,
              limit,
              returned: items.length,
              hasMore: nextCursor !== null,
              nextCursor,
            },
            coverage: {
              readableSourceCount: sources.length,
              readableSourceCounts: sourceCounts,
              unsupportedContainerCount: unsupported.length,
              unsupportedContainers: unsupported.slice(0, 50),
              unsupportedContainerListTruncated: unsupported.length > 50,
              scanComplete: nextCursor === null && unsupported.length === 0,
            },
            nextAction:
              nextCursor !== null
                ? "Continue with read_slide_translatable_texts using page.nextCursor before reaching an audit conclusion."
                : unsupported.length > 0
                  ? "Do not claim a complete translation audit until the unsupported containers are inspected or explicitly excluded."
                  : "This slide's readable translatable content has been fully scanned.",
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
        error instanceof Error
          ? error.message
          : "Failed to read slide translatable texts";
      return toolError(message);
    }
  },
});
