import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import {
  attachSlideResultMetadata,
  unpackSlideZipResult,
} from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";

function getShapeIdentity(shape: Element): { id: string; name: string } {
  const cNvPr = shape.getElementsByTagNameNS(NS_P, "cNvPr")[0];
  return {
    id: cNvPr?.getAttribute("id") ?? "",
    name: cNvPr?.getAttribute("name") ?? "",
  };
}

function getOmittedShapeType(
  shape: Element,
): "table" | "chart" | "group" | "graphicFrame" {
  if (shape.localName === "grpSp") return "group";
  const graphicData = shape.getElementsByTagNameNS(NS_A, "graphicData")[0];
  const uri = graphicData?.getAttribute("uri") ?? "";
  if (uri.includes("/table")) return "table";
  if (uri.includes("/chart")) return "chart";
  return "graphicFrame";
}

function getShapeText(shape: Element): string | null {
  const txBody = shape.getElementsByTagNameNS(NS_P, "txBody")[0];
  if (!txBody) return null;
  const paragraphs = Array.from(txBody.childNodes).filter(
    (node): node is Element =>
      node.nodeType === 1 &&
      (node as Element).localName === "p" &&
      (node as Element).namespaceURI === NS_A,
  );
  return paragraphs
    .map((paragraph) => {
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
    })
    .join("\n");
}

export const readSlideTextsTool = defineTool({
  name: "read_slide_texts",
  label: "Read Slide Texts",
  description:
    "Read all text shapes on one slide as a compact, paginated text index. " +
    "Returns shape IDs/names/text only and deliberately omits fonts, colors, " +
    "geometry, and raw OOXML. Covers normal text shapes; tables/charts/groups " +
    "require a specialized or OOXML read. Prefer this for reading or translation.",
  parameters: Type.Object({
    ...slideTargetParameterProperties,
    offset: Type.Optional(
      Type.Number({ description: "0-based text-shape offset. Default 0." }),
    ),
    limit: Type.Optional(
      Type.Number({ description: "Maximum text shapes. Default 50, max 100." }),
    ),
    max_bytes: Type.Optional(
      Type.Number({
        description:
          "Maximum text budget before JSON wrapping. Default and max 8000 bytes.",
      }),
    ),
    include_empty: Type.Optional(
      Type.Boolean({
        description: "Include empty text shapes. Default false.",
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
          const shapes = Array.from(doc.getElementsByTagNameNS(NS_P, "sp"));
          const allItems = shapes
            .map((shape) => ({
              ...getShapeIdentity(shape),
              text: getShapeText(shape),
            }))
            .filter(
              (item): item is { id: string; name: string; text: string } =>
                item.text !== null &&
                (params.include_empty === true || item.text.length > 0),
            );
          const omittedShapes = [
            ...Array.from(doc.getElementsByTagNameNS(NS_P, "graphicFrame")),
            ...Array.from(doc.getElementsByTagNameNS(NS_P, "grpSp")),
          ].map((shape) => ({
            ...getShapeIdentity(shape),
            type: getOmittedShapeType(shape),
          }));
          const omittedShapeCounts = omittedShapes.reduce<
            Record<string, number>
          >((counts, shape) => {
            counts[shape.type] = (counts[shape.type] ?? 0) + 1;
            return counts;
          }, {});

          const offset = Math.max(0, Math.floor(params.offset ?? 0));
          const limit = Math.min(
            100,
            Math.max(1, Math.floor(params.limit ?? 50)),
          );
          const maxBytes = Math.min(
            8_000,
            Math.max(1_000, Math.floor(params.max_bytes ?? 8_000)),
          );
          const encoder = new TextEncoder();
          const items: Array<{
            id: string;
            name: string;
            text: string;
            textOmitted?: boolean;
          }> = [];
          const oversizedShapeIds: string[] = [];
          let usedBytes = 0;

          for (
            let index = offset;
            index < allItems.length && items.length < limit;
            index++
          ) {
            const item = allItems[index];
            const fixedBytes =
              encoder.encode(item.id + item.name).byteLength + 80;
            const available = maxBytes - usedBytes - fixedBytes;
            if (available <= 0) break;

            const textBytes = encoder.encode(item.text);
            if (textBytes.byteLength > available) {
              items.push({ ...item, text: "", textOmitted: true });
              oversizedShapeIds.push(item.id);
              usedBytes = maxBytes;
              break;
            }
            items.push(item);
            usedBytes += fixedBytes + textBytes.byteLength;
          }

          const nextOffset =
            offset + items.length < allItems.length
              ? offset + items.length
              : null;
          const nextActions: string[] = [];
          if (oversizedShapeIds.length > 0) {
            nextActions.push(
              "Use read_slide_text with format=plain and paragraph pagination for each oversized shape ID.",
            );
          }
          if (omittedShapes.length > 0) {
            nextActions.push(
              "This compact read does not cover table/chart/group containers. Inspect the listed IDs with list_slide_shapes and use a specialized Office.js/OOXML path before claiming the slide is fully translated.",
            );
          }
          return {
            schemaVersion: 3,
            items,
            page: {
              offset,
              limit,
              total: allItems.length,
              returned: items.length,
              hasMore: nextOffset !== null,
              nextOffset,
            },
            omittedFields: [
              "font",
              "fontSize",
              "color",
              "position",
              "geometry",
              "rawOoxml",
            ],
            omittedShapeTypes: Object.keys(omittedShapeCounts),
            omittedShapeCounts,
            omittedShapeCount: omittedShapes.length,
            omittedShapes: omittedShapes.slice(0, 50),
            omittedShapeListTruncated: omittedShapes.length > 50,
            oversizedShapeIds,
            nextAction: nextActions.length > 0 ? nextActions.join(" ") : null,
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
        error instanceof Error ? error.message : "Failed to read slide texts";
      return toolError(message);
    }
  },
});
