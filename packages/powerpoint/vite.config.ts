import { svelte } from "@sveltejs/vite-plugin-svelte";
import tailwindcss from "@tailwindcss/postcss";
import autoprefixer from "autoprefixer";
import { createRequire } from "module";
import path from "path";
import { fileURLToPath } from "url";
import { defineConfig } from "vite";
import { nodePolyfills } from "vite-plugin-node-polyfills";
import { viteStaticCopy } from "vite-plugin-static-copy";
import {
  createCcSwitchProxy,
  isAllowedCcSwitchRequest,
} from "./cc-switch-proxy";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const pkg = require("./package.json");

async function getHttpsOptions() {
  try {
    const devCerts = await import("office-addin-dev-certs");
    const certs = await devCerts.getHttpsServerOptions();
    return { ca: certs.ca, key: certs.key, cert: certs.cert };
  } catch {
    console.warn("Could not load office-addin-dev-certs, HTTPS disabled");
    return undefined;
  }
}

export default defineConfig(async ({ mode }) => {
  const dev = mode === "development";
  const urlDev = "https://localhost:3001/";
  const urlProd = "https://klaragraff.github.io/pptxmate/";

  return {
    base: dev ? "/" : "/pptxmate/",
    root: "src",
    publicDir: "../public",

    build: {
      outDir: "../dist",
      emptyOutDir: true,
      sourcemap: true,
      rollupOptions: {
        input: {
          taskpane: path.resolve(__dirname, "src/taskpane.html"),
          commands: path.resolve(__dirname, "src/commands.html"),
        },
      },
    },

    resolve: {
      alias: {
        "node:util/types": path.resolve(
          __dirname,
          "src/shims/util-types-shim.js",
        ),
      },
    },

    define: {
      "process.env": JSON.stringify({}),
      "process.versions": "undefined",
      "process.browser": JSON.stringify(true),
      __APP_VERSION__: JSON.stringify(pkg.version),
    },

    css: {
      postcss: {
        plugins: [tailwindcss(), autoprefixer()],
      },
    },

    plugins: [
      svelte(),

      {
        name: "pptxmate-cc-switch-request-guard",
        configureServer(server) {
          server.middlewares.use((request, response, next) => {
            const pathname = (request.url ?? "").split("?", 1)[0];
            if (pathname !== "/v1" && !pathname.startsWith("/v1/")) {
              next();
              return;
            }
            if (isAllowedCcSwitchRequest(request.headers)) {
              next();
              return;
            }
            response.statusCode = 403;
            response.setHeader("Content-Type", "text/plain; charset=utf-8");
            response.end("Forbidden CC Switch proxy origin.");
          });
        },
      },

      nodePolyfills({
        include: [
          "buffer",
          "stream",
          "util",
          "url",
          "http",
          "https",
          "zlib",
          "path",
          "os",
          "assert",
          "events",
          "querystring",
          "punycode",
          "string_decoder",
          "constants",
          "vm",
          "process",
        ],
        globals: {
          Buffer: true,
          process: true,
        },
      }),

      viteStaticCopy({
        targets: [
          {
            src: "../../../LICENSE",
            dest: ".",
          },
          {
            src: "../../../THIRD_PARTY_NOTICES.md",
            dest: ".",
          },
          {
            src: "../manifest*.xml",
            dest: ".",
            transform: {
              encoding: "utf8",
              handler(content: string) {
                if (dev) return content;
                return content.replace(new RegExp(urlDev, "g"), urlProd);
              },
            },
          },
        ],
      }),
    ],

    server: {
      https: await getHttpsOptions(),
      port: 3001,
      cors: false,
      proxy: createCcSwitchProxy(),
    },
  };
});
