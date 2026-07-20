import type { AgentContext } from "@office-agents/core";
import { sandboxedEval } from "@office-agents/core";
import { Type } from "@sinclair/typebox";
import {
  cleanupSlideMasters,
  type MasterCleanupReceipt,
} from "../pptx/master-cleanup";
import {
  loadSlideDirectory,
  type SlideTargetReference,
} from "../pptx/slide-directory";
import { safeRun, withSlideZip } from "../pptx/slide-zip";
import { escapeXml } from "../pptx/xml-utils";
import { unpackSlideZipResult } from "./slide-target-result";
import { defineTool, toolSuccess } from "./types";

/* global PowerPoint */

function buildMasterFailure(
  error: unknown,
  metadata: Record<string, unknown> = {},
  knownMutationCompleted = false,
) {
  const record =
    error && typeof error === "object"
      ? (error as Record<string, unknown>)
      : {};
  const mutationCompleted =
    knownMutationCompleted || record.mutationCompleted === true;
  const failure: Record<string, unknown> = {
    success: false,
    ...metadata,
    mutationCompleted,
    mutationState: mutationCompleted ? "uncertain" : "not_started",
    error: {
      code:
        typeof record.code === "string"
          ? record.code
          : "EDIT_SLIDE_MASTER_FAILED",
      message: error instanceof Error ? error.message : String(error),
    },
    mutation: {
      kind: "master",
      status: mutationCompleted ? "uncertain" : "not_started",
      mutationCompleted,
    },
  };
  for (const [source, destination] of [
    ["cleanupPhase", "cleanupPhase"],
    ["expectedVersion", "expectedDirectoryVersion"],
    ["currentVersion", "currentDirectoryVersion"],
    ["temporarySlideId", "temporarySlideId"],
  ] as const) {
    if (typeof record[source] === "string") {
      failure[destination] = record[source];
    }
  }
  return failure;
}

export function createEditSlideMasterTool(ctx: AgentContext) {
  return defineTool({
    name: "edit_slide_master",
    label: "Edit Slide Master",
    description:
      "Edit slide master and layouts via OOXML — set backgrounds, decorative elements, " +
      "fonts, theme colors, and placeholders. Use this for any visual element that should " +
      "appear on all slides.",
    parameters: Type.Object({
      code: Type.String({
        description:
          "Async function body receiving { zip, markDirty }. zip is a JSZip archive " +
          "containing the full PPTX structure including ppt/slideMasters/ and ppt/slideLayouts/. " +
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
        const operation = await safeRun(async (context) => {
          const directory = await loadSlideDirectory(context);
          if (directory.slideIds.length === 0) {
            throw new Error(
              "The presentation has no slide to carry the master edit.",
            );
          }
          const target: SlideTargetReference = {
            slide_id: directory.slideIds[0],
            directory_version: directory.directoryVersion,
          };
          const zipValue = await withSlideZip(context, target, async (args) => {
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
          try {
            const cleanup = await cleanupSlideMasters(
              context,
              zipValue.directoryVersion,
            );
            return { zipValue, target, cleanup, cleanupError: null };
          } catch (cleanupError) {
            return {
              zipValue,
              target,
              cleanup: null,
              cleanupError,
            };
          }
        });
        const { result, metadata } = unpackSlideZipResult(
          operation.zipValue,
          operation.target,
        );
        const metadataRecord = metadata as Record<string, unknown>;
        const primaryMutationCompleted = metadata.replacementSlideId !== null;
        if (operation.cleanupError) {
          return toolSuccess(
            buildMasterFailure(
              operation.cleanupError,
              metadataRecord,
              primaryMutationCompleted,
            ),
          );
        }

        const cleanup = operation.cleanup as MasterCleanupReceipt;
        const cleanupMutationCompleted =
          cleanup.reassignedSlideCount > 0 ||
          cleanup.temporarySlidesCreated > 0;
        const mutationCompleted =
          primaryMutationCompleted || cleanupMutationCompleted;

        return toolSuccess({
          success: true,
          ...metadata,
          mutationCompleted,
          mutationState: mutationCompleted ? "completed" : "not_started",
          mutation: {
            kind: "master",
            status: mutationCompleted ? "completed" : "not_started",
            mutationCompleted,
            cleanup,
          },
          cleanup,
          result: result !== undefined ? result : null,
        });
      } catch (error) {
        return toolSuccess(buildMasterFailure(error));
      }
    },
  });
}
