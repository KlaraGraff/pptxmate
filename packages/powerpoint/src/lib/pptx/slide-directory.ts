import { Type } from "@sinclair/typebox";

/* global PowerPoint */

export interface SlideTargetReference {
  /** Stable PowerPoint slide ID. Preferred over slide_index. */
  slide_id?: string;
  /** Current zero-based position. Kept for backwards compatibility only. */
  slide_index?: number;
  /** Version returned by a lightweight slide-directory read. */
  directory_version?: string;
}

export interface SlideDirectorySnapshot {
  slideIds: string[];
  directoryVersion: string;
  indexById: ReadonlyMap<string, number>;
}

export interface ResolvedSlideTarget {
  slideId: string;
  slideIndex: number;
  directoryVersion: string;
  /** True when the caller supplied a version that no longer matches. */
  directoryChanged: boolean;
  /** True when the caller's optional numeric hint no longer matches the ID. */
  indexMismatch: boolean;
  /** True when the target was resolved from the legacy numeric index. */
  usedLegacyIndex: boolean;
}

export type SlideMutationState = "not_started" | "completed" | "uncertain";

export interface SlideMutationError extends Error {
  readonly code: string;
  readonly mutationCompleted: boolean;
  readonly mutationState: Exclude<SlideMutationState, "completed">;
}

function errorMessage(error: unknown, fallbackMessage: string): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  if (typeof error === "string" && error.trim().length > 0) return error;
  return fallbackMessage;
}

function errorCode(error: unknown, fallbackCode: string): string {
  if (typeof error !== "object" || error === null) return fallbackCode;
  const code = (error as { code?: unknown }).code;
  return typeof code === "string" && code.length > 0 ? code : fallbackCode;
}

function hasMutationState(error: unknown): error is SlideMutationError {
  if (!(error instanceof Error)) return false;
  const mutationState = (error as Partial<SlideMutationError>).mutationState;
  return mutationState === "not_started" || mutationState === "uncertain";
}

const PRESERVED_ERROR_METADATA_KEYS = [
  "expectedVersion",
  "currentVersion",
  "directoryVersion",
  "slideId",
  "originalSlideId",
  "replacementSlideId",
  "newSlideId",
] as const;

function copyErrorMetadata(source: unknown, target: Error): void {
  if (typeof source !== "object" || source === null) return;
  const record = source as Record<string, unknown>;
  for (const key of PRESERVED_ERROR_METADATA_KEYS) {
    if (record[key] === undefined || key in target) continue;
    Object.defineProperty(target, key, {
      configurable: true,
      enumerable: true,
      value: record[key],
    });
  }
}

export class SlideMutationNotStartedError
  extends Error
  implements SlideMutationError
{
  readonly mutationCompleted = false;
  readonly mutationState = "not_started" as const;

  constructor(
    message: string,
    readonly code = "SLIDE_MUTATION_NOT_STARTED",
  ) {
    super(message);
    this.name = "SlideMutationNotStartedError";
  }
}

export class SlideMutationUncertainError
  extends Error
  implements SlideMutationError
{
  readonly mutationCompleted = true;
  readonly mutationState = "uncertain" as const;

  constructor(
    message: string,
    readonly code = "SLIDE_MUTATION_UNCERTAIN",
  ) {
    super(message);
    this.name = "SlideMutationUncertainError";
  }
}

/** Normalize any failure before the host write sync into a non-replay-blocking result. */
export function toSlideMutationNotStartedError(
  error: unknown,
  fallbackMessage = "Slide mutation did not start.",
  fallbackCode = "SLIDE_MUTATION_NOT_STARTED",
): SlideMutationError {
  if (hasMutationState(error)) return error;
  const normalized = new SlideMutationNotStartedError(
    errorMessage(error, fallbackMessage),
    errorCode(error, fallbackCode),
  );
  copyErrorMetadata(error, normalized);
  return normalized;
}

/** Normalize any failure after the host write sync starts into a no-replay result. */
export function toSlideMutationUncertainError(
  error: unknown,
  fallbackMessage = "The slide write may have completed; verify before retrying.",
  fallbackCode = "SLIDE_MUTATION_UNCERTAIN",
): SlideMutationError {
  if (hasMutationState(error) && error.mutationState === "uncertain") {
    return error;
  }
  const normalized = new SlideMutationUncertainError(
    errorMessage(error, fallbackMessage),
    errorCode(error, fallbackCode),
  );
  copyErrorMetadata(error, normalized);
  return normalized;
}

export const slideTargetParameterProperties = {
  slide_id: Type.Optional(
    Type.String({
      description:
        "Stable slide ID from list_slides. Preferred for all reads and writes.",
      minLength: 1,
      maxLength: 256,
    }),
  ),
  slide_index: Type.Optional(
    Type.Integer({
      description:
        "Current zero-based slide position. Compatibility fallback only; do not treat it as a stable ID.",
      minimum: 0,
    }),
  ),
  directory_version: Type.Optional(
    Type.String({
      description:
        "Directory version returned by list_slides. Stable IDs are re-resolved if it changed; index-only calls fail closed.",
      minLength: 1,
      maxLength: 80,
    }),
  ),
} as const;

export const slideTargetParameters = Type.Object(
  slideTargetParameterProperties,
);

export function toSlideTargetReference(
  input: Partial<SlideTargetReference>,
): SlideTargetReference {
  return {
    ...(input.slide_id === undefined ? {} : { slide_id: input.slide_id }),
    ...(input.slide_index === undefined
      ? {}
      : { slide_index: input.slide_index }),
    ...(input.directory_version === undefined
      ? {}
      : { directory_version: input.directory_version }),
  };
}

export class SlideDirectoryChangedError extends SlideMutationNotStartedError {
  readonly code = "SLIDE_DIRECTORY_CHANGED";

  constructor(
    readonly expectedVersion: string,
    readonly currentVersion: string,
  ) {
    super(
      `[SLIDE_DIRECTORY_CHANGED] Slide directory changed (expected ${expectedVersion}, current ${currentVersion}). Refresh list_slides and resolve the target by slide_id before writing.`,
      "SLIDE_DIRECTORY_CHANGED",
    );
    this.name = "SlideDirectoryChangedError";
  }
}

/** The host accepted a replacement, but the post-write directory was not the expected one. */
export class SlideDirectoryChangedDuringWriteError extends SlideMutationUncertainError {
  readonly code = "SLIDE_DIRECTORY_CHANGED_DURING_WRITE";

  constructor(
    readonly expectedVersion: string,
    readonly currentVersion: string,
  ) {
    super(
      `[SLIDE_DIRECTORY_CHANGED_DURING_WRITE mutationState=uncertain mutationCompleted=true] Slide directory changed during the write (expected ${expectedVersion}, current ${currentVersion}). The write may have completed; do not replay it. Refresh list_slides and verify the affected slide first.`,
      "SLIDE_DIRECTORY_CHANGED_DURING_WRITE",
    );
    this.name = "SlideDirectoryChangedDuringWriteError";
  }
}

export class SlideTargetNotFoundError extends SlideMutationNotStartedError {
  readonly code = "SLIDE_TARGET_NOT_FOUND";

  constructor(readonly slideId: string) {
    super(
      `[SLIDE_TARGET_NOT_FOUND] Slide ID "${slideId}" is no longer present. Refresh list_slides instead of falling back to an old slide_index.`,
      "SLIDE_TARGET_NOT_FOUND",
    );
    this.name = "SlideTargetNotFoundError";
  }
}

/**
 * Compact deterministic fingerprint of slide membership and order.
 * Selection and slide content are deliberately excluded.
 */
export function getSlideDirectoryVersion(slideIds: readonly string[]): string {
  const input = `${slideIds.length}\u0000${slideIds.join("\u0000")}`;
  let hash = 2166136261;
  for (const byte of new TextEncoder().encode(input)) {
    hash ^= byte;
    hash = Math.imul(hash, 16777619);
  }
  return `directory-v1:fnv1a32:${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

export function createSlideDirectorySnapshot(
  slides: readonly Pick<PowerPoint.Slide, "id">[],
): SlideDirectorySnapshot {
  const slideIds = slides.map((slide) => slide.id);
  return {
    slideIds,
    directoryVersion: getSlideDirectoryVersion(slideIds),
    indexById: new Map(slideIds.map((slideId, index) => [slideId, index])),
  };
}

export async function loadSlideDirectory(
  context: PowerPoint.RequestContext,
): Promise<SlideDirectorySnapshot> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();
  return createSlideDirectorySnapshot(slides.items);
}

export function assertSlideDirectoryVersion(
  snapshot: SlideDirectorySnapshot,
  expectedVersion: string | undefined,
): void {
  if (
    expectedVersion !== undefined &&
    expectedVersion !== snapshot.directoryVersion
  ) {
    throw new SlideDirectoryChangedError(
      expectedVersion,
      snapshot.directoryVersion,
    );
  }
}

export function resolveSlideTarget(
  snapshot: SlideDirectorySnapshot,
  target: SlideTargetReference,
): ResolvedSlideTarget {
  if (target.slide_id !== undefined) {
    const slideIndex = snapshot.indexById.get(target.slide_id);
    if (slideIndex === undefined) {
      throw new SlideTargetNotFoundError(target.slide_id);
    }
    const hintedIndex = target.slide_index;
    return {
      slideId: target.slide_id,
      slideIndex,
      directoryVersion: snapshot.directoryVersion,
      directoryChanged:
        target.directory_version !== undefined &&
        target.directory_version !== snapshot.directoryVersion,
      indexMismatch: hintedIndex !== undefined && hintedIndex !== slideIndex,
      usedLegacyIndex: false,
    };
  }

  // An index has no stable meaning after a structural change. Fail closed
  // before any read with side effects or mutation when the caller has a stale
  // directory token and did not provide an authoritative ID.
  assertSlideDirectoryVersion(snapshot, target.directory_version);

  if (
    target.slide_index === undefined ||
    !Number.isInteger(target.slide_index) ||
    target.slide_index < 0 ||
    target.slide_index >= snapshot.slideIds.length
  ) {
    throw new SlideMutationNotStartedError(
      `Slide target requires slide_id or a valid slide_index (0-${Math.max(0, snapshot.slideIds.length - 1)}).`,
      "SLIDE_TARGET_INVALID",
    );
  }

  return {
    slideId: snapshot.slideIds[target.slide_index],
    slideIndex: target.slide_index,
    directoryVersion: snapshot.directoryVersion,
    directoryChanged: false,
    indexMismatch: false,
    usedLegacyIndex: true,
  };
}

export function resolveSlideIds(
  snapshot: SlideDirectorySnapshot,
  requestedIds: readonly string[] | undefined,
  requestedIndices: readonly number[] | undefined,
): {
  ids: string[];
  indices: number[];
  usedLegacyIndices: boolean;
  relocatedIds: string[];
} {
  if (requestedIds && requestedIds.length > 0) {
    const ids = Array.from(new Set(requestedIds));
    const unknown = ids.filter((id) => !snapshot.indexById.has(id));
    if (unknown.length > 0) {
      throw new SlideTargetNotFoundError(unknown[0]);
    }
    const indices = ids.map((id) => snapshot.indexById.get(id) as number);
    const relocatedIds = ids.filter((id, uniqueIndex) => {
      const originalIndex = requestedIds.indexOf(id);
      const hint = requestedIndices?.[originalIndex];
      return hint !== undefined && Math.floor(hint) !== indices[uniqueIndex];
    });
    return {
      ids,
      indices,
      usedLegacyIndices: false,
      relocatedIds,
    };
  }

  if (requestedIndices && requestedIndices.length > 0) {
    const indices = Array.from(
      new Set(requestedIndices.map((index) => Math.floor(index))),
    );
    const invalid = indices.filter(
      (index) => index < 0 || index >= snapshot.slideIds.length,
    );
    if (invalid.length > 0) {
      throw new SlideMutationNotStartedError(
        `Slide indices out of range: ${invalid.join(", ")} (slide count: ${snapshot.slideIds.length}). Call list_slides to refresh the directory.`,
        "SLIDE_TARGET_INVALID",
      );
    }
    return {
      ids: indices.map((index) => snapshot.slideIds[index]),
      indices,
      usedLegacyIndices: true,
      relocatedIds: [],
    };
  }

  return {
    ids: [...snapshot.slideIds],
    indices: snapshot.slideIds.map((_id, index) => index),
    usedLegacyIndices: false,
    relocatedIds: [],
  };
}
