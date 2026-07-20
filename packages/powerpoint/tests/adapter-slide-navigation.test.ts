import path from "node:path";
import { fileURLToPath } from "node:url";
import { createServer, type Plugin } from "vite";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  run: vi.fn(),
  setSelectedSlides: vi.fn(),
  loadSlides: vi.fn(),
  sync: vi.fn(async () => undefined),
}));

const packageRoot = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "..",
);

function adapterDependencyStubs(): Plugin {
  const modules = new Map<string, string>([
    [
      "@office-agents/core",
      "export const getOrCreateDocumentId = async () => 'test-document';",
    ],
    [
      "./recovery-router",
      "export const getPowerPointToolRecoveryInfo = () => ({}); export const normalizePowerPointToolArgsForReplay = (_name, args) => args;",
    ],
    [
      "./request-router",
      "export const compactPowerPointContext = (messages) => messages; export const routePowerPointMessage = () => 'general';",
    ],
    [
      "./system-prompt",
      "export const buildPowerPointSystemPrompt = () => 'prompt';",
    ],
    [
      "./tool-router",
      "export const needsPowerPointThemeContext = () => false; export const powerPointToolAllowlist = () => null; export const routePowerPointMessage = () => 'general';",
    ],
    ["./tools", "export const createPptTools = () => [];"],
    ["./vfs/custom-commands", "export const getCustomCommands = () => [];"],
  ]);
  const virtualPrefix = "\0adapter-test-stub:";

  return {
    name: "adapter-test-dependency-stubs",
    enforce: "pre",
    resolveId(source) {
      if (modules.has(source)) return `${virtualPrefix}${source}`;
      if (source.endsWith("selection-indicator.svelte")) {
        return `${virtualPrefix}selection-indicator`;
      }
      if (source.endsWith("powerpoint-officejs-api.d.ts?raw")) {
        return `${virtualPrefix}powerpoint-api-dts`;
      }
      return null;
    },
    load(id) {
      if (!id.startsWith(virtualPrefix)) return null;
      const key = id.slice(virtualPrefix.length);
      if (key === "selection-indicator") return "export default {};";
      if (key === "powerpoint-api-dts") return "export default '';";
      return modules.get(key) ?? null;
    },
  };
}

beforeEach(() => {
  vi.clearAllMocks();
  mocks.run.mockImplementation(async (callback: (context: object) => unknown) =>
    callback({
      presentation: {
        setSelectedSlides: mocks.setSelectedSlides,
        slides: {
          load: mocks.loadSlides,
          items: [{ id: "wrong-slide-from-index" }],
        },
      },
      sync: mocks.sync,
    }),
  );
  vi.stubGlobal("PowerPoint", { run: mocks.run });
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("PowerPoint adapter slide navigation", () => {
  it("prefers the stable modified slide ID when an index is also returned", async () => {
    const server = await createServer({
      root: packageRoot,
      configFile: false,
      envFile: false,
      appType: "custom",
      logLevel: "silent",
      server: { middlewareMode: true },
      define: { __APP_VERSION__: JSON.stringify("test") },
      plugins: [adapterDependencyStubs()],
    });

    try {
      const { createPowerPointAdapter } = await server.ssrLoadModule(
        "/src/lib/adapter.ts",
      );
      const adapter = createPowerPointAdapter();

      adapter.onToolResult?.(
        "tool-1",
        JSON.stringify({
          success: true,
          _modifiedSlideId: "original-H-replacement",
          _modifiedSlide: 6,
        }),
        false,
      );

      await vi.waitFor(() => {
        expect(mocks.setSelectedSlides).toHaveBeenCalledWith([
          "original-H-replacement",
        ]);
      });
      expect(mocks.loadSlides).not.toHaveBeenCalled();
      expect(mocks.sync).toHaveBeenCalledTimes(1);
    } finally {
      await server.close();
    }
  });
});
