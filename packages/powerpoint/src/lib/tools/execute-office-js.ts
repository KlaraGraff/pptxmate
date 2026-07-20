import type { AgentContext } from "@office-agents/core";
import { sandboxedEval } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import {
  loadSlideDirectory,
  type ResolvedSlideTarget,
  resolveSlideTarget,
  slideTargetParameterProperties,
  toSlideMutationNotStartedError,
  toSlideMutationUncertainError,
  toSlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun } from "../pptx/slide-zip";
import { defineTool, toolError, toolSuccess } from "./types";

/* global PowerPoint */

const POSITIONAL_SLIDE_ACCESS_RE =
  /\b(?:context\s*\.\s*)?presentation\s*\.\s*slides\s*\.\s*getItemAt\s*\(/;
const TARGET_SLIDE_REFERENCE_RE = /\btargetSlide\b/;

function officeErrorMessage(error: unknown): string {
  if (
    typeof OfficeExtension !== "undefined" &&
    error instanceof OfficeExtension.Error
  ) {
    const parts = [error.message];
    if (error.code) parts.push(`Code: ${error.code}`);
    if (error.debugInfo) {
      const { errorLocation, statement, surroundingStatements } =
        error.debugInfo;
      if (errorLocation) parts.push(`Location: ${errorLocation}`);
      if (statement) parts.push(`Statement: ${statement}`);
      if (surroundingStatements?.length) {
        parts.push(`Context: ${surroundingStatements.join("; ")}`);
      }
    }
    return parts.join("\n");
  }
  return error instanceof Error
    ? error.message
    : "Unknown error executing code";
}

function targetResult(target: ResolvedSlideTarget | undefined) {
  if (!target) return {};
  return {
    slideId: target.slideId,
    slideIndex: target.slideIndex,
    positionOneIndexed: target.slideIndex + 1,
    directoryVersion: target.directoryVersion,
    directoryChanged: target.directoryChanged,
    relocated: target.indexMismatch || target.directoryChanged,
    usedLegacyIndex: target.usedLegacyIndex,
  };
}

export function createExecuteOfficeJsTool(ctx: AgentContext) {
  return defineTool({
    name: "execute_office_js",
    label: "Execute Office.js Code",
    description:
      "Execute Office.js JavaScript code to interact with the PowerPoint document. " +
      "The code receives context and an optional stable targetSlide, and runs inside PowerPoint.run(). " +
      "For any existing-slide operation, pass slide_id from list_slides and use targetSlide instead of positional slide access. " +
      "Use this for any document operations like adding slides, shapes, text, and formatting.",
    parameters: Type.Object({
      ...slideTargetParameterProperties,
      code: Type.String({
        description:
          "Async function body that receives 'context: PowerPoint.RequestContext'. " +
          "When slide_id or slide_index is supplied it also receives targetSlide, targetSlideId, targetSlideIndex, and directoryVersion. " +
          "Use targetSlide for existing-slide work; context.presentation.slides.getItemAt(...) is rejected because its identity becomes stale after deletion or reordering. " +
          "Must call context.sync() to execute batched operations and load() to read properties. " +
          "Return JSON-serializable results. " +
          "readFile(path) returns Promise<string> and readFileBuffer(path) returns Promise<Uint8Array> " +
          "to read files from the virtual filesystem. " +
          "writeFile(path, content) returns Promise<void> to write string or Uint8Array to the virtual filesystem. " +
          "btoa(string) and atob(base64) are available for base64 encoding/decoding.",
      }),
      explanation: Type.Optional(
        Type.String({
          description:
            "Brief explanation of what this code does (max 100 chars)",
          maxLength: 100,
        }),
      ),
    }),
    execute: async (_toolCallId, params) => {
      let executionStarted = false;
      let resolvedTarget: ResolvedSlideTarget | undefined;
      try {
        const hasTarget =
          params.slide_id !== undefined || params.slide_index !== undefined;
        if (POSITIONAL_SLIDE_ACCESS_RE.test(params.code)) {
          throw Object.assign(
            new Error(
              "[POSITIONAL_SLIDE_ACCESS_REJECTED] Bind the page with list_slides, pass slide_id, and use targetSlide instead of presentation.slides.getItemAt(...).",
            ),
            { code: "POSITIONAL_SLIDE_ACCESS_REJECTED" },
          );
        }

        if (!hasTarget && TARGET_SLIDE_REFERENCE_RE.test(params.code)) {
          throw Object.assign(
            new Error(
              "[SLIDE_TARGET_REQUIRED] Code that uses targetSlide must include slide_id or slide_index.",
            ),
            { code: "SLIDE_TARGET_REQUIRED" },
          );
        }
        if (!hasTarget && params.directory_version !== undefined) {
          throw Object.assign(
            new Error(
              "[SLIDE_TARGET_REQUIRED] directory_version requires slide_id or slide_index.",
            ),
            { code: "SLIDE_TARGET_REQUIRED" },
          );
        }

        const result = await safeRun(async (context) => {
          let targetSlide: PowerPoint.Slide | undefined;
          if (hasTarget) {
            const directory = await loadSlideDirectory(context);
            resolvedTarget = resolveSlideTarget(
              directory,
              toSlideTargetReference(params),
            );
            targetSlide = context.presentation.slides.getItem(
              resolvedTarget.slideId,
            );
          }

          executionStarted = true;
          return sandboxedEval(params.code, {
            context,
            targetSlide,
            targetSlideId: resolvedTarget?.slideId,
            targetSlideIndex: resolvedTarget?.slideIndex,
            directoryVersion: resolvedTarget?.directoryVersion,
            PowerPoint,
            Office,
            readFile: (path: string) => ctx.readFile(path),
            readFileBuffer: (path: string) => ctx.readFileBuffer(path),
            writeFile: (path: string, content: string | Uint8Array) =>
              ctx.writeFile(path, content),
          });
        });

        return toolSuccess({
          success: true,
          result: result ?? null,
          ...targetResult(resolvedTarget),
          mutationCompleted: true,
          mutationState: "completed",
        });
      } catch (error) {
        const message = officeErrorMessage(error);
        const normalized = executionStarted
          ? toSlideMutationUncertainError(error, message)
          : toSlideMutationNotStartedError(error, message);
        return toolError(normalized.message, {
          ...normalized,
          ...targetResult(resolvedTarget),
        });
      }
    },
  });
}
