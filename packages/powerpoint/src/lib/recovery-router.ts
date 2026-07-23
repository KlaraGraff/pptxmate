import type { ToolRecoveryInfo } from "@office-agents/core";

const READ_TOOLS = new Set([
  "read",
  "list_slides",
  "read_slides",
  "screenshot_slide",
  "list_slide_shapes",
  "read_slide_texts",
  "read_slide_translatable_texts",
  "read_slide_text",
  "verify_slides",
]);

const MUTATION_TOOLS = new Set([
  "execute_office_js",
  "edit_slide_text",
  "patch_slide_text",
  "update_slide_text",
  "edit_slide_xml",
  "edit_slide_chart",
  "edit_slide_master",
  "duplicate_slide",
]);

const MUTATION_KINDS = {
  execute_office_js: "arbitrary",
  edit_slide_text: "text",
  patch_slide_text: "text",
  update_slide_text: "text",
  edit_slide_chart: "arbitrary",
  edit_slide_master: "arbitrary",
  duplicate_slide: "structure",
  edit_slide_xml: "arbitrary",
} as const;

const VERIFICATION_KINDS = {
  list_slides: ["structure"],
  read_slides: ["text"],
  read_slide_texts: ["text"],
  read_slide_translatable_texts: ["text"],
  read_slide_text: ["text"],
  verify_slides: ["layout"],
  screenshot_slide: ["layout"],
} as const;

type RecoveryRecord = Record<string, unknown>;

const BASH_SLIDE_WRITE_RE = /\binsert-(?:image|icon)\b/i;

function firstString(
  record: RecoveryRecord,
  fields: readonly string[],
): string | undefined {
  for (const field of fields) {
    const value = record[field];
    if (typeof value === "string" && value.trim().length > 0) return value;
  }
  return undefined;
}

function firstIndex(
  record: RecoveryRecord,
  fields: readonly string[],
): number | undefined {
  for (const field of fields) {
    const value = record[field];
    if (
      typeof value === "number" &&
      Number.isSafeInteger(value) &&
      value >= 0
    ) {
      return value;
    }
  }
  return undefined;
}

function firstStringArray(
  record: RecoveryRecord,
  fields: readonly string[],
): string[] | undefined {
  for (const field of fields) {
    const value = record[field];
    if (!Array.isArray(value)) continue;
    const strings = value.filter(
      (item): item is string => typeof item === "string" && item.length > 0,
    );
    if (strings.length > 0) return strings;
  }
  return undefined;
}

function firstIndexArray(
  record: RecoveryRecord,
  fields: readonly string[],
): number[] | undefined {
  for (const field of fields) {
    const value = record[field];
    if (!Array.isArray(value)) continue;
    const indices = value.filter(
      (item): item is number =>
        typeof item === "number" && Number.isSafeInteger(item) && item >= 0,
    );
    if (indices.length > 0) return indices;
  }
  return undefined;
}

function bashSlideWriteScope(record: RecoveryRecord): RecoveryRecord | null {
  const operation = firstString(record, ["operation", "commandName"]);
  const command = firstString(record, ["command"]);
  if (
    operation !== "insert-image" &&
    operation !== "insert-icon" &&
    (!command || !BASH_SLIDE_WRITE_RE.test(command))
  ) {
    return null;
  }

  const scope: RecoveryRecord = {};
  const resultSlideId = firstString(record, [
    "_modifiedSlideId",
    "replacementSlideId",
    "newSlideId",
    "slideId",
    "slide_id",
  ]);
  const commandSlideId = command?.match(
    /--slide(?:-|_)id=(?:"([^"]+)"|'([^']+)'|([^\s;&|]+))/i,
  );
  const slideId =
    resultSlideId ??
    commandSlideId?.[1] ??
    commandSlideId?.[2] ??
    commandSlideId?.[3];
  if (slideId) {
    scope.slide_id = slideId;
  } else if (command) {
    const positional = command.match(
      /\binsert-(?:image|icon)\s+(?:"[^"]*"|'[^']*'|\S+)\s+(\d+)/i,
    );
    if (positional) scope.slide_index = Number.parseInt(positional[1], 10) - 1;
  }

  const directoryMatch = command?.match(
    /--directory(?:-|_)version=(?:"([^"]+)"|'([^']+)'|([^\s;&|]+))/i,
  );
  const directoryVersion =
    firstString(record, ["directoryVersion", "directory_version"]) ??
    directoryMatch?.[1] ??
    directoryMatch?.[2] ??
    directoryMatch?.[3];
  if (directoryVersion) scope.directory_version = directoryVersion;
  return scope;
}

function normalizeBashSlideWrite(
  record: RecoveryRecord,
  overrideSlideId?: string,
): RecoveryRecord | undefined {
  if (
    typeof record.command !== "string" ||
    !BASH_SLIDE_WRITE_RE.test(record.command)
  ) {
    return undefined;
  }
  const idMatch = record.command.match(
    /--slide(?:-|_)id=(?:"([^"]+)"|'([^']+)'|([^\s;&|]+))/i,
  );
  const slideId =
    overrideSlideId ?? idMatch?.[1] ?? idMatch?.[2] ?? idMatch?.[3];
  let command = record.command
    .replace(/--slide(?:-|_)id=(?:"[^"]+"|'[^']+'|[^\s;&|]+)/gi, "")
    .replace(/--directory(?:-|_)version=(?:"[^"]+"|'[^']+'|[^\s;&|]+)/gi, "")
    .replace(/\s+/g, " ")
    .trim();
  if (slideId) command = `${command} --slide-id=${slideId}`;
  return { command };
}

export function getPowerPointToolRecoveryInfo(
  toolName: string,
  args: unknown,
): ToolRecoveryInfo {
  const record =
    args && typeof args === "object" ? (args as Record<string, unknown>) : {};
  if (toolName === "bash") {
    const bashScope = bashSlideWriteScope(record);
    return bashScope
      ? {
          effect: "write",
          mutationKind: "layout",
          verificationKinds: [],
          scope:
            Object.keys(bashScope).length > 0
              ? (bashScope as Record<string, string | number | boolean>)
              : undefined,
        }
      : { effect: "read", verificationKinds: [] };
  }
  const scope: Record<string, string | number | boolean> = {};
  // A returned modified ID is the post-write object. Otherwise slide_id is
  // authoritative across reordering/deletion of other slides. Keep a numeric
  // position only when no stable ID exists, so recovery never targets a stale
  // occupant after a structural change.
  const slideId = firstString(record, [
    "_modifiedSlideId",
    "modifiedSlideId",
    "replacement_slide_id",
    "replacementSlideId",
    "new_slide_id",
    "newSlideId",
    "slide_id",
    "slideId",
    "current_slide_id",
    "currentSlideId",
  ]);
  if (slideId) {
    scope.slide_id = slideId;
  } else {
    const slideIndex = firstIndex(record, [
      "_modifiedSlide",
      "slide_index",
      "slideIndex",
      "current_slide_index",
      "currentSlideIndex",
    ]);
    if (slideIndex !== undefined) scope.slide_index = slideIndex;
  }

  const directoryVersion = firstString(record, [
    "directory_version",
    "directoryVersion",
  ]);
  if (directoryVersion) scope.directory_version = directoryVersion;

  const shapeId = firstString(record, ["shape_id", "shapeId"]);
  if (shapeId) scope.shape_id = shapeId;

  if (toolName === "update_slide_text" && Array.isArray(record.updates)) {
    const shapeIds = record.updates
      .flatMap((update) =>
        update && typeof update === "object" && !Array.isArray(update)
          ? [
              (update as Record<string, unknown>).shape_id ??
                (update as Record<string, unknown>).shapeId,
            ]
          : [],
      )
      .filter((shapeId): shapeId is string => typeof shapeId === "string")
      .slice(0, 50);
    if (shapeIds.length > 0) scope.shape_ids = shapeIds.join(",");
  }

  const slideIds = firstStringArray(record, ["slide_ids", "slideIds"]);
  if (slideIds) {
    scope.slide_ids = slideIds.slice(0, 4).join(",");
  } else {
    const slideIndices = firstIndexArray(record, [
      "slide_indices",
      "slideIndices",
    ]);
    if (slideIndices) {
      scope.slide_indices = slideIndices.slice(0, 12).join(",");
    }
  }

  const originalSlideId = firstString(record, [
    "original_slide_id",
    "originalSlideId",
  ]);
  if (originalSlideId) scope.original_slide_id = originalSlideId;

  const replacementSlideId = firstString(record, [
    "replacement_slide_id",
    "replacementSlideId",
  ]);
  if (replacementSlideId) scope.replacement_slide_id = replacementSlideId;

  const sourceSlideId = firstString(record, [
    "source_slide_id",
    "sourceSlideId",
  ]);
  if (sourceSlideId) scope.source_slide_id = sourceSlideId;

  const newSlideId = firstString(record, ["new_slide_id", "newSlideId"]);
  if (newSlideId) scope.new_slide_id = newSlideId;

  return {
    effect: READ_TOOLS.has(toolName)
      ? "read"
      : MUTATION_TOOLS.has(toolName)
        ? "write"
        : "unknown",
    mutationKind:
      MUTATION_KINDS[toolName as keyof typeof MUTATION_KINDS] ?? "arbitrary",
    verificationKinds: [
      ...(VERIFICATION_KINDS[toolName as keyof typeof VERIFICATION_KINDS] ??
        []),
    ],
    scope: Object.keys(scope).length > 0 ? scope : undefined,
  };
}

export function normalizePowerPointToolArgsForReplay(
  toolName: string,
  args: unknown,
): unknown {
  if (!args || typeof args !== "object" || Array.isArray(args)) return args;
  const normalized = { ...(args as Record<string, unknown>) };
  delete normalized.explanation;

  const stableSlideId = firstString(normalized, [
    "_modifiedSlideId",
    "modifiedSlideId",
    "replacement_slide_id",
    "replacementSlideId",
    "new_slide_id",
    "newSlideId",
    "slide_id",
    "slideId",
    "current_slide_id",
    "currentSlideId",
  ]);
  const stableSlideIds = firstStringArray(normalized, [
    "slide_ids",
    "slideIds",
  ]);
  if (toolName === "bash") {
    const normalizedCommand = normalizeBashSlideWrite(
      normalized,
      stableSlideId,
    );
    if (normalizedCommand) return normalizedCommand;
  }
  if (stableSlideId) {
    normalized.slide_id = stableSlideId;
    delete normalized._modifiedSlideId;
    delete normalized.modifiedSlideId;
    delete normalized.replacement_slide_id;
    delete normalized.replacementSlideId;
    delete normalized.new_slide_id;
    delete normalized.newSlideId;
    delete normalized.slideId;
    delete normalized.current_slide_id;
    delete normalized.currentSlideId;
    delete normalized.slide_index;
    delete normalized.slideIndex;
    delete normalized.current_slide_index;
    delete normalized.currentSlideIndex;
    delete normalized._modifiedSlide;
    // The directory token and numeric position are optimistic snapshot guards,
    // not part of the intended mutation identity when a stable ID is present.
    delete normalized.directory_version;
    delete normalized.directoryVersion;
  } else if (stableSlideIds) {
    normalized.slide_ids = stableSlideIds;
    delete normalized.slideIds;
    delete normalized.slide_indices;
    delete normalized.slideIndices;
    delete normalized.directory_version;
    delete normalized.directoryVersion;
  }

  if (toolName === "edit_slide_text" || toolName === "patch_slide_text") {
    delete normalized.expected_text;
    delete normalized.expected_text_hash;
    if (typeof normalized.text === "string") {
      normalized.mode = normalized.mode ?? "replace";
    } else {
      delete normalized.mode;
    }
  }
  if (toolName === "update_slide_text" && Array.isArray(normalized.updates)) {
    normalized.updates = normalized.updates.map((update) => {
      if (!update || typeof update !== "object" || Array.isArray(update)) {
        return update;
      }
      const normalizedUpdate = { ...(update as Record<string, unknown>) };
      delete normalizedUpdate.expected_text;
      delete normalizedUpdate.expected_text_hash;
      normalizedUpdate.mode = normalizedUpdate.mode ?? "replace";
      return normalizedUpdate;
    });
  }
  return normalized;
}
