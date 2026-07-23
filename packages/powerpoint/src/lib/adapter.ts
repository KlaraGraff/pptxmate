import type {
  AppAdapter,
  DocumentMetadataRequest,
  MessagePreparationInfo,
  SkillMeta,
} from "@office-agents/core";
import { getOrCreateDocumentId } from "@office-agents/core";
import SelectionIndicator from "./components/selection-indicator.svelte";
import pptApiDts from "./docs/powerpoint-officejs-api.d.ts?raw";
import { getSlideDirectoryVersion } from "./pptx/slide-directory";
import {
  getPowerPointToolRecoveryInfo,
  normalizePowerPointToolArgsForReplay,
} from "./recovery-router";
import {
  type CompactContextMessage,
  compactPowerPointContext,
  type PowerPointTaskRoute,
} from "./request-router";
import { buildPowerPointSystemPrompt } from "./system-prompt";
import {
  needsPowerPointThemeContext,
  powerPointToolAllowlist,
  routePowerPointMessage,
} from "./tool-router";
import { createPptTools } from "./tools";
import { getCustomCommands } from "./vfs/custom-commands";

/* global PowerPoint, Office */

const STORAGE_NAMESPACE = {
  dbName: "OpenPPTDB_v1",
  dbVersion: 1,
  localStoragePrefix: "openppt",
  documentSettingsPrefix: "openppt",
  documentIdSettingsKey: "openppt-presentation-id",
};

function selectMessageTools(
  userMessage: string,
  ctx: Parameters<typeof createPptTools>[0],
  info: MessagePreparationInfo,
) {
  const tools = createPptTools(ctx);
  const allow = powerPointToolAllowlist(userMessage, info);
  if (!allow) return tools;
  return tools.filter((tool) => allow.has(tool.name));
}

export function createPowerPointAdapter(): AppAdapter {
  return {
    tools: (ctx) => createPptTools(ctx),
    toolsForMessage: selectMessageTools,
    customCommands: getCustomCommands,
    hasImageSearch: true,
    staticFiles: {
      "/home/user/docs/powerpoint-officejs-api.d.ts": pptApiDts,
    },

    appName: "PPTXMate",
    metadataTag: "ppt_context",
    storageNamespace: STORAGE_NAMESPACE,
    appVersion: __APP_VERSION__,
    emptyStateMessage:
      "Start a conversation to create or edit your presentation",
    SelectionIndicator,
    buildSystemPrompt: buildPowerPointSystemPrompt,
    buildSystemPromptForMessage: (
      userMessage: string,
      skills: SkillMeta[],
      commandSnippets: string[],
      info?: MessagePreparationInfo,
    ) =>
      buildPowerPointSystemPrompt(
        skills,
        commandSnippets,
        routePowerPointMessage(userMessage, info),
      ),
    metadataHistory: "latest",
    getToolRecoveryInfo: getPowerPointToolRecoveryInfo,
    normalizeToolArgsForReplay: normalizePowerPointToolArgsForReplay,
    toolExecution: "sequential",
    transformContext: async (messages, _signal, info) => {
      const baseMaxChars = info
        ? Math.min(
            60_000,
            Math.max(
              12_000,
              info.contextWindow * 2 - info.systemPromptChars - 12_000,
            ),
          )
        : 60_000;
      const maxChars = Math.max(
        6_000,
        Math.floor(baseMaxChars / 2 ** Math.max(0, info?.recoveryAttempt ?? 0)),
      );
      return compactPowerPointContext(
        messages as unknown as CompactContextMessage[],
        { maxChars },
      ) as unknown as typeof messages;
    },

    getDocumentId: async () => {
      return getOrCreateDocumentId(STORAGE_NAMESPACE);
    },

    getDocumentMetadata: async (request?: DocumentMetadataRequest) => {
      try {
        const route = request
          ? routePowerPointMessage(request.userMessage, request.info)
          : "general";
        const includeThemeContext = needsPowerPointThemeContext(
          request?.userMessage ?? "",
          route,
        );
        const metadata = await getPresentationMetadata(
          route,
          includeThemeContext,
        );
        return { metadata };
      } catch {
        return null;
      }
    },

    onToolResult: (_toolCallId, result, isError) => {
      if (isError) return;
      try {
        const parsed = JSON.parse(result);
        if (typeof parsed._modifiedSlideId === "string") {
          navigateToSlideById(parsed._modifiedSlideId).catch(console.error);
        } else if (typeof parsed._modifiedSlide === "number") {
          navigateToSlideByIndex(parsed._modifiedSlide).catch(console.error);
        }
      } catch {
        // Not JSON or no modified slide info
      }
    },
  };
}

async function navigateToSlideById(slideId: string): Promise<void> {
  return PowerPoint.run(async (context) => {
    context.presentation.setSelectedSlides([slideId]);
    await context.sync();
  });
}

/** Legacy recovery path for tools that still return an index during migration. */
async function navigateToSlideByIndex(slideIndex: number): Promise<void> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");
    await context.sync();
    if (slideIndex >= 0 && slideIndex < slides.items.length) {
      context.presentation.setSelectedSlides([slides.items[slideIndex].id]);
      await context.sync();
    }
  });
}

const DEFAULT_OFFICE_THEME_COLORS: Record<string, string> = {
  Dark1: "#000000",
  Dark2: "#44546A",
  Light1: "#FFFFFF",
  Light2: "#E7E6E6",
  Accent1: "#4472C4",
  Accent2: "#ED7D31",
  Accent3: "#A5A5A5",
  Accent4: "#FFC000",
  Accent5: "#5B9BD5",
  Accent6: "#70AD47",
};

const THEME_COLOR_KEYS = [
  "Dark1",
  "Dark2",
  "Light1",
  "Light2",
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
] as const;

function normalizeColor(c: string): string {
  return c.replace(/^#/, "").toUpperCase();
}

async function detectThemeDefault(
  master: PowerPoint.SlideMaster,
  context: PowerPoint.RequestContext,
): Promise<{ isDefault: boolean; confidence: "high" | "medium" | "low" }> {
  try {
    const scheme = master.themeColorScheme;
    const colorResults: Record<
      string,
      OfficeExtension.ClientResult<string>
    > = {};

    for (const key of THEME_COLOR_KEYS) {
      colorResults[key] = scheme.getThemeColor(key);
    }

    master.shapes.load("items/id");
    await context.sync();

    let matchCount = 0;
    for (const key of THEME_COLOR_KEYS) {
      const actual = normalizeColor(colorResults[key].value);
      const expected = normalizeColor(DEFAULT_OFFICE_THEME_COLORS[key]);
      if (actual === expected) matchCount++;
    }

    const total = THEME_COLOR_KEYS.length;
    const masterShapeCount = master.shapes.items.length;

    if (matchCount === total) {
      return { isDefault: true, confidence: "high" };
    }
    if (matchCount >= total - 2) {
      return {
        isDefault: masterShapeCount <= 2,
        confidence: "medium",
      };
    }
    return {
      isDefault: false,
      confidence: matchCount >= 3 ? "medium" : "high",
    };
  } catch {
    return { isDefault: true, confidence: "low" };
  }
}

async function getPresentationMetadata(
  route: PowerPointTaskRoute = "general",
  includeThemeContext = route === "create",
): Promise<object> {
  return PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items/id");

    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/id");

    let selectedShapesCollection: PowerPoint.ShapeScopedCollection | undefined;
    const needsGeometry = route === "layout" || route === "verify";
    try {
      selectedShapesCollection = context.presentation.getSelectedShapes();
      selectedShapesCollection.load(
        needsGeometry
          ? "items/name,items/type,items/id,items/left,items/top,items/width,items/height"
          : "items/name,items/type,items/id",
      );
    } catch {
      // getSelectedShapes may not be available on older hosts
    }

    await context.sync();

    const idToIndex = new Map(slides.items.map((s, i) => [s.id, i]));
    const selectedIndices = selectedSlides.items.map((s) => ({
      slideId: s.id,
      positionOneIndexed: (idToIndex.get(s.id) ?? 0) + 1,
    }));

    const selectedShapes = (selectedShapesCollection?.items ?? []).map((s) => {
      const base = { name: s.name, type: s.type, id: s.id };
      if (!needsGeometry) return base;
      return {
        ...base,
        left: s.left,
        top: s.top,
        width: s.width,
        height: s.height,
      };
    });

    const metadata: Record<string, unknown> = {
      schemaVersion: 2,
      route,
      slideCount: slides.items.length,
      directoryVersion: getSlideDirectoryVersion(
        slides.items.map((slide) => slide.id),
      ),
      selectedSlides: selectedIndices,
      selectedShapes,
      omittedFields:
        route === "text" || route === "translationAudit" || route === "general"
          ? [
              "slideGeometry",
              "shapeGeometry",
              "theme",
              "masters",
              "layouts",
              "fonts",
              "colors",
              "slideText",
            ]
          : route === "verify"
            ? ["theme", "masters", "layouts", "fonts", "colors", "slideText"]
            : includeThemeContext
              ? ["slideText", "shapeFormatting"]
              : ["slideText", "shapeFormatting", "theme", "masters", "layouts"],
    };

    // Geometry/theme/master data is opt-in. Text-only requests should not
    // force PowerPoint to hydrate or serialize presentation styling.
    if (route === "layout" || route === "create" || route === "verify") {
      const pageSetup = context.presentation.pageSetup;
      pageSetup.load(["slideWidth", "slideHeight"]);
      await context.sync();
      metadata.slideWidth = pageSetup.slideWidth;
      metadata.slideHeight = pageSetup.slideHeight;
    }

    if (includeThemeContext) {
      const masters = context.presentation.slideMasters;
      masters.load("items/id");
      await context.sync();

      for (const master of masters.items) {
        master.layouts.load("items/name,items/id");
      }
      await context.sync();

      metadata.masters = masters.items.map((m, mi) => ({
        index: mi,
        layouts: m.layouts.items.map((l) => ({ name: l.name, id: l.id })),
      }));

      const themeResult =
        masters.items.length > 0
          ? await detectThemeDefault(masters.items[0], context)
          : { isDefault: true, confidence: "low" as const };
      metadata.isDefaultTheme = themeResult.isDefault;
      metadata.themeDetectionConfidence = themeResult.confidence;
    }

    if (route === "create") {
      // Only load lightweight shape IDs when a blank-vs-existing decision is
      // actually needed. The IDs never leave the host; only the boolean does.
      const shapeCollections = slides.items.map((slide) => {
        const shapes = slide.shapes;
        shapes.load("items/id");
        return shapes;
      });
      await context.sync();
      metadata.hasContent = shapeCollections.some(
        (shapes) => shapes.items.length > 0,
      );
    }

    return metadata;
  });
}
