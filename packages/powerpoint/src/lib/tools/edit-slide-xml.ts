import type { AgentContext } from "@office-agents/core";
import { sandboxedEval } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import {
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import { escapeXml } from "../pptx/xml-utils";
import { unpackSlideZipResult } from "./slide-target-result";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

export function createEditSlideXmlTool(ctx: AgentContext) {
  return defineTool({
    name: "edit_slide_xml",
    label: "Edit Slide XML",
    description:
      "Edit raw OOXML of a PowerPoint slide. Use for advanced formatting, " +
      "custom XML manipulation, diagrams, or anything not covered by other tools. " +
      "Use escapeXml(text) to escape special characters when embedding text in XML.",
    parameters: Type.Object({
      ...slideTargetParameterProperties,
      code: Type.String({
        description:
          "Async function body receiving { zip, markDirty }. zip is a JSZip archive of the slide. " +
          "Call markDirty() if you modified files. " +
          "Globals: escapeXml(text) for safe XML text embedding, " +
          "readFile(path) returns Promise<string> and readFileBuffer(path) returns Promise<Uint8Array> " +
          "to read files from the virtual filesystem (e.g. uploaded images, SVGs). " +
          "writeFile(path, content) returns Promise<void> to write string or Uint8Array to the virtual filesystem.",
      }),
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
        const zipValue = await safeRun(async (context) => {
          return withSlideZip(context, target, async (args) => {
            return sandboxedEval(params.code, {
              ...args,
              escapeXml,
              readFile: (path: string) => ctx.readFile(path),
              readFileBuffer: (path: string) => ctx.readFileBuffer(path),
              writeFile: (path: string, content: string | Uint8Array) =>
                ctx.writeFile(path, content),
              DOMParser,
              XMLSerializer,
            });
          });
        });
        const { result, metadata } = unpackSlideZipResult(zipValue, target);

        return toolSuccess({
          success: true,
          ...metadata,
          result: result !== undefined ? result : null,
        });
      } catch (error) {
        const normalized = toSlideMutationNotStartedError(
          error,
          "Failed to edit slide XML",
        );
        return toolError(normalized.message, normalized);
      }
    },
    modifiedSlide: (params, result) =>
      result && typeof result.slideId === "string"
        ? result.slideId
        : (params.slide_id ?? params.slide_index),
  });
}
