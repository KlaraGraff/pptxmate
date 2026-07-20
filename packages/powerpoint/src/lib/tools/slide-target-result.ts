import type { SlideTargetReference } from "../pptx/slide-directory";
import type { SlideZipResult } from "../pptx/slide-zip";

export interface SlideResultMetadata {
  originalSlideId?: string;
  slideId?: string;
  replacementSlideId?: string | null;
  slideIndex?: number;
  positionOneIndexed?: number;
  directoryVersion?: string;
  directoryChanged?: boolean;
  inputDirectoryChanged?: boolean;
  relocated?: boolean;
  usedLegacyIndex?: boolean;
  mutationCompleted?: boolean;
  mutationState?: "not_started" | "completed";
}

export function unpackSlideZipResult<T>(
  value: SlideZipResult<T> | T,
  _target: SlideTargetReference,
): { result: T; metadata: SlideResultMetadata } {
  if (
    typeof value === "object" &&
    value !== null &&
    "result" in value &&
    "slideId" in value &&
    "directoryVersion" in value
  ) {
    const wrapped = value as SlideZipResult<T>;
    return {
      result: wrapped.result,
      metadata: {
        originalSlideId: wrapped.originalSlideId,
        slideId: wrapped.slideId,
        replacementSlideId: wrapped.replacementSlideId,
        slideIndex: wrapped.slideIndex,
        positionOneIndexed: wrapped.slideIndex + 1,
        directoryVersion: wrapped.directoryVersion,
        directoryChanged: wrapped.directoryChanged,
        inputDirectoryChanged: wrapped.inputDirectoryChanged,
        relocated: wrapped.relocated,
        usedLegacyIndex: wrapped.usedLegacyIndex,
        mutationCompleted: wrapped.mutationCompleted,
        mutationState: wrapped.mutationState,
      },
    };
  }

  return {
    result: value as T,
    metadata: {},
  };
}

export function attachSlideResultMetadata<T>(
  result: T,
  metadata: SlideResultMetadata,
): T | (Record<string, unknown> & SlideResultMetadata) {
  if (typeof result !== "object" || result === null || Array.isArray(result)) {
    return result;
  }
  return { ...(result as Record<string, unknown>), ...metadata };
}
