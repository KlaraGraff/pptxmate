import {
  type CustomCommandsResult,
  type DescribedCommand,
  getSharedCustomCommands,
  type StorageNamespace,
} from "@office-agents/core";
import { defineCommand } from "just-bash/browser";
import type { SlideTargetReference } from "../pptx/slide-directory";
import { type SlideZipResult, safeRun, withSlideZip } from "../pptx/slide-zip";

async function resolveVfsPath(
  ctx: { cwd: string; fs: { readFileBuffer(p: string): Promise<Uint8Array> } },
  filePath: string,
): Promise<Uint8Array> {
  const resolved = filePath.startsWith("/")
    ? filePath
    : `${ctx.cwd}/${filePath}`;
  return ctx.fs.readFileBuffer(resolved);
}

/* global PowerPoint, Office */

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_R =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const EMU_PER_INCH = 914400;
const EMU_PER_CM = 360000;
const EMU_PER_POINT = 12700;
const DEFAULT_IMAGE_BOX_PT = { width: 360, height: 270 };
const DEFAULT_ICON_BOX_PT = { width: 72, height: 72 };

function parseUnit(value: number, unit: string): number {
  switch (unit) {
    case "in":
      return Math.round(value * EMU_PER_INCH);
    case "cm":
      return Math.round(value * EMU_PER_CM);
    case "emu":
      return Math.round(value);
    case "pt":
      return Math.round(value * EMU_PER_POINT);
    default:
      return Math.round(value * EMU_PER_INCH);
  }
}

function pointsToUnit(value: number, unit: string): number {
  switch (unit) {
    case "in":
      return value / 72;
    case "cm":
      return (value * 2.54) / 72;
    case "emu":
      return value * EMU_PER_POINT;
    default:
      return value;
  }
}

function parseSvgNumber(value: string | null): number | undefined {
  if (!value) return undefined;
  const match = value.trim().match(/^([+-]?(?:\d+\.?\d*|\.\d+))/);
  if (!match) return undefined;
  const parsed = Number.parseFloat(match[1]);
  return Number.isFinite(parsed) ? parsed : undefined;
}

async function getImageDimensions(
  data: Uint8Array,
  ext: string,
): Promise<{ width: number; height: number } | undefined> {
  if (ext === "svg") {
    const svgText = new TextDecoder().decode(data);
    const doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
    const svg = doc.documentElement;
    if (svg?.localName !== "svg") return undefined;

    const viewBox = svg.getAttribute("viewBox");
    if (viewBox) {
      const parts = viewBox
        .trim()
        .split(/[\s,]+/)
        .map((part) => Number.parseFloat(part));
      if (
        parts.length === 4 &&
        parts.every((part) => Number.isFinite(part)) &&
        parts[2] > 0 &&
        parts[3] > 0
      ) {
        return { width: parts[2], height: parts[3] };
      }
    }

    const width = parseSvgNumber(svg.getAttribute("width"));
    const height = parseSvgNumber(svg.getAttribute("height"));
    if (width && height && width > 0 && height > 0) {
      return { width, height };
    }

    return undefined;
  }

  const bytes = new Uint8Array(data.byteLength);
  bytes.set(data);
  const blob = new Blob([bytes.buffer]);
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    if (img.naturalWidth > 0 && img.naturalHeight > 0) {
      return { width: img.naturalWidth, height: img.naturalHeight };
    }
    return undefined;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function resolveImageSize({
  intrinsic,
  requestedWidth,
  requestedHeight,
  widthProvided,
  heightProvided,
  defaultWidth,
  defaultHeight,
}: {
  intrinsic?: { width: number; height: number };
  requestedWidth?: number;
  requestedHeight?: number;
  widthProvided: boolean;
  heightProvided: boolean;
  defaultWidth: number;
  defaultHeight: number;
}): { width: number; height: number } {
  const fallbackWidth = widthProvided
    ? (requestedWidth ?? defaultWidth)
    : defaultWidth;
  const fallbackHeight = heightProvided
    ? (requestedHeight ?? defaultHeight)
    : defaultHeight;

  if (!intrinsic || intrinsic.width <= 0 || intrinsic.height <= 0) {
    return { width: fallbackWidth, height: fallbackHeight };
  }

  const aspect = intrinsic.width / intrinsic.height;

  if (widthProvided && heightProvided) {
    const boxWidth = requestedWidth ?? defaultWidth;
    const boxHeight = requestedHeight ?? defaultHeight;
    if (boxWidth / boxHeight > aspect) {
      return { width: boxHeight * aspect, height: boxHeight };
    }
    return { width: boxWidth, height: boxWidth / aspect };
  }

  if (widthProvided) {
    const width = requestedWidth ?? defaultWidth;
    return { width, height: width / aspect };
  }

  if (heightProvided) {
    const height = requestedHeight ?? defaultHeight;
    return { width: height * aspect, height };
  }

  if (defaultWidth / defaultHeight > aspect) {
    return { width: defaultHeight * aspect, height: defaultHeight };
  }

  return { width: defaultWidth, height: defaultWidth / aspect };
}

function detectMimeType(
  path: string,
  data: Uint8Array,
): { mime: string; ext: string } {
  if (
    data[0] === 0x89 &&
    data[1] === 0x50 &&
    data[2] === 0x4e &&
    data[3] === 0x47
  )
    return { mime: "image/png", ext: "png" };
  if (data[0] === 0xff && data[1] === 0xd8)
    return { mime: "image/jpeg", ext: "jpeg" };
  if (data[0] === 0x47 && data[1] === 0x49 && data[2] === 0x46)
    return { mime: "image/gif", ext: "gif" };
  if (data[0] === 0x42 && data[1] === 0x4d)
    return { mime: "image/bmp", ext: "bmp" };

  const fileExt = path.split(".").pop()?.toLowerCase() || "";
  const map: Record<string, { mime: string; ext: string }> = {
    png: { mime: "image/png", ext: "png" },
    jpg: { mime: "image/jpeg", ext: "jpeg" },
    jpeg: { mime: "image/jpeg", ext: "jpeg" },
    gif: { mime: "image/gif", ext: "gif" },
    svg: { mime: "image/svg+xml", ext: "svg" },
    webp: { mime: "image/webp", ext: "webp" },
    bmp: { mime: "image/bmp", ext: "bmp" },
    tiff: { mime: "image/tiff", ext: "tiff" },
    tif: { mime: "image/tiff", ext: "tiff" },
  };
  return map[fileExt] || { mime: "image/png", ext: "png" };
}

const SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}";
const NS_ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";

async function rasterizeSvgToPng(
  svgData: Uint8Array,
  widthPx: number,
  heightPx: number,
): Promise<Uint8Array> {
  const svgText = new TextDecoder().decode(svgData);
  const blob = new Blob([svgText], { type: "image/svg+xml" });
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();

    const canvas = document.createElement("canvas");
    canvas.width = widthPx;
    canvas.height = heightPx;
    const ctx = canvas.getContext("2d");
    if (!ctx) throw new Error("Failed to create canvas 2D context");
    ctx.drawImage(img, 0, 0, widthPx, heightPx);

    const pngData = await new Promise<Uint8Array>((resolve, reject) => {
      canvas.toBlob((b) => {
        if (!b) return reject(new Error("Canvas toBlob failed"));
        b.arrayBuffer().then((buf) => resolve(new Uint8Array(buf)));
      }, "image/png");
    });

    canvas.width = 0;
    canvas.height = 0;
    return pngData;
  } finally {
    URL.revokeObjectURL(url);
  }
}

interface InsertImageParams {
  target: SlideTargetReference;
  data: Uint8Array;
  mime: string;
  ext: string;
  shapeName: string;
  offX: number;
  offY: number;
  cx: number;
  cy: number;
  mediaPrefix?: string;
}

interface InsertedShapeResult {
  shapeId: string;
  shapeName: string;
}

type InsertImageZipResult = SlideZipResult<InsertedShapeResult>;

class CommandInputError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message);
    this.name = "CommandInputError";
  }
}

function parseCommandSlideTarget(
  slideArg: string | undefined,
  flags: Record<string, string>,
): { target: SlideTargetReference; requestedPositionOneIndexed?: number } {
  const slideId = flags["slide-id"] ?? flags.slide_id;
  const directoryVersion =
    flags["directory-version"] ?? flags.directory_version;
  let requestedPositionOneIndexed: number | undefined;
  if (slideArg !== undefined) {
    if (!/^\d+$/.test(slideArg)) {
      throw new CommandInputError(
        "INVALID_SLIDE_TARGET",
        "Slide must be a positive number (1-based).",
      );
    }
    requestedPositionOneIndexed = Number.parseInt(slideArg, 10);
    if (requestedPositionOneIndexed < 1) {
      throw new CommandInputError(
        "INVALID_SLIDE_TARGET",
        "Slide must be a positive number (1-based).",
      );
    }
  }
  if (!slideId && requestedPositionOneIndexed === undefined) {
    throw new CommandInputError(
      "SLIDE_TARGET_REQUIRED",
      "Provide --slide-id=<ID> from list_slides or a legacy 1-based slide number.",
    );
  }
  return {
    target: {
      ...(slideId ? { slide_id: slideId } : {}),
      ...(requestedPositionOneIndexed === undefined
        ? {}
        : { slide_index: requestedPositionOneIndexed - 1 }),
      ...(directoryVersion ? { directory_version: directoryVersion } : {}),
    },
    requestedPositionOneIndexed,
  };
}

function structuredCommandFailure(
  operation: "insert-image" | "insert-icon",
  error: unknown,
  fallbackCode: string,
) {
  const record =
    error && typeof error === "object"
      ? (error as Record<string, unknown>)
      : {};
  const message = error instanceof Error ? error.message : String(error);
  const code = typeof record.code === "string" ? record.code : fallbackCode;
  const mutationCompleted = record.mutationCompleted === true;
  const receipt: Record<string, unknown> = {
    success: false,
    operation,
    mutationCompleted,
    mutationState: mutationCompleted ? "uncertain" : "not_started",
    error: { code, message },
    mutation: {
      kind: "slide-content",
      status: mutationCompleted ? "uncertain" : "not_started",
      mutationCompleted,
    },
  };
  for (const [source, destination] of [
    ["expectedVersion", "expectedDirectoryVersion"],
    ["currentVersion", "currentDirectoryVersion"],
    ["temporarySlideId", "temporarySlideId"],
  ] as const) {
    if (typeof record[source] === "string") {
      receipt[destination] = record[source];
    }
  }
  return {
    stdout: JSON.stringify(receipt),
    stderr: "",
    exitCode: 1,
  };
}

function buildInsertReceipt(
  operation: "insert-image" | "insert-icon",
  zipResult: InsertImageZipResult,
  details: Record<string, unknown>,
): Record<string, unknown> {
  const mutationCompleted = zipResult.replacementSlideId !== null;
  return {
    success: true,
    operation,
    mutationCompleted,
    mutationState: mutationCompleted ? "completed" : "not_started",
    originalSlideId: zipResult.originalSlideId,
    slideId: zipResult.slideId,
    replacementSlideId: zipResult.replacementSlideId,
    slideIndex: zipResult.slideIndex,
    positionOneIndexed: zipResult.slideIndex + 1,
    directoryVersion: zipResult.directoryVersion,
    directoryChanged: zipResult.directoryChanged,
    inputDirectoryChanged: zipResult.inputDirectoryChanged,
    relocated: zipResult.relocated,
    usedLegacyIndex: zipResult.usedLegacyIndex,
    shapeId: zipResult.result.shapeId,
    shapeName: zipResult.result.shapeName,
    mutation: {
      kind: "slide-content",
      status: mutationCompleted ? "completed" : "not_applied",
      mutationCompleted,
      replacedSlide: mutationCompleted,
    },
    ...details,
  };
}

async function insertImageIntoSlide(
  params: InsertImageParams,
): Promise<InsertImageZipResult> {
  const { target, data, ext, mime, shapeName, offX, offY, cx, cy } = params;
  const prefix =
    params.mediaPrefix || `vfs_image_${Date.now()}_${crypto.randomUUID()}`;
  const isSvg = ext === "svg";
  let pngFallback: Uint8Array | undefined;
  if (isSvg) {
    const DPI = 96;
    const pxW = Math.round((cx / EMU_PER_INCH) * DPI * 2);
    const pxH = Math.round((cy / EMU_PER_INCH) * DPI * 2);
    pngFallback = await rasterizeSvgToPng(data, pxW, pxH);
  }

  return safeRun(async (context) => {
    return withSlideZip(context, target, async ({ zip, markDirty }) => {
      const relsPath = "ppt/slides/_rels/slide1.xml.rels";
      const relsFile = zip.file(relsPath);
      let relsXml: string;
      if (relsFile) {
        relsXml = await relsFile.async("string");
      } else {
        relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_RELS}"></Relationships>`;
      }

      const rIdMatches = [...relsXml.matchAll(/Id="rId(\d+)"/g)];
      let nextId =
        rIdMatches.reduce(
          (max, m) => Math.max(max, Number.parseInt(m[1], 10)),
          0,
        ) + 1;

      let blipRId: string;
      let svgRId: string | undefined;

      if (isSvg && pngFallback) {
        const pngPath = `ppt/media/${prefix}_fallback.png`;
        zip.file(pngPath, pngFallback);
        blipRId = `rId${nextId++}`;
        const pngRel = `<Relationship Id="${blipRId}" Type="${IMAGE_REL_TYPE}" Target="../media/${prefix}_fallback.png"/>`;
        relsXml = relsXml.replace(
          "</Relationships>",
          `${pngRel}</Relationships>`,
        );

        const svgPath = `ppt/media/${prefix}.svg`;
        zip.file(svgPath, data);
        svgRId = `rId${nextId++}`;
        const svgRel = `<Relationship Id="${svgRId}" Type="${IMAGE_REL_TYPE}" Target="../media/${prefix}.svg"/>`;
        relsXml = relsXml.replace(
          "</Relationships>",
          `${svgRel}</Relationships>`,
        );
      } else {
        const imagePath = `ppt/media/${prefix}.${ext}`;
        zip.file(imagePath, data);
        blipRId = `rId${nextId++}`;
        const rel = `<Relationship Id="${blipRId}" Type="${IMAGE_REL_TYPE}" Target="../media/${prefix}.${ext}"/>`;
        relsXml = relsXml.replace("</Relationships>", `${rel}</Relationships>`);
      }

      zip.file(relsPath, relsXml);

      const slideFile = zip.file("ppt/slides/slide1.xml");
      if (!slideFile) throw new Error("Slide XML not found in archive");

      const slideXml = await slideFile.async("string");
      const doc = new DOMParser().parseFromString(slideXml, "text/xml");

      const spTree = doc.getElementsByTagNameNS(NS_P, "spTree")[0];
      if (!spTree) throw new Error("Shape tree not found in slide XML");

      const existingIds = new Set<number>();
      const allEls = doc.getElementsByTagName("*");
      for (let i = 0; i < allEls.length; i++) {
        const el = allEls[i];
        if (el.localName === "cNvPr") {
          const id = el.getAttribute("id");
          if (id) existingIds.add(Number.parseInt(id, 10));
        }
      }
      let shapeId = 1;
      while (existingIds.has(shapeId)) shapeId++;

      const escapedName = shapeName
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

      let blipInner = "";
      if (isSvg && svgRId) {
        blipInner = `
      <a:extLst>
        <a:ext uri="${SVG_EXT_URI}">
          <asvg:svgBlip xmlns:asvg="${NS_ASVG}" r:embed="${svgRId}"/>
        </a:ext>
      </a:extLst>`;
      }

      const picXml = `<p:pic xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${escapedName}"/>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${blipRId}">${blipInner}
    </a:blip>
    <a:stretch>
      <a:fillRect/>
    </a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${offX}" y="${offY}"/>
      <a:ext cx="${cx}" cy="${cy}"/>
    </a:xfrm>
    <a:prstGeom prst="rect">
      <a:avLst/>
    </a:prstGeom>
  </p:spPr>
</p:pic>`;

      const picDoc = new DOMParser().parseFromString(picXml, "text/xml");
      const parseError = picDoc.getElementsByTagName("parsererror")[0];
      if (parseError) {
        throw new Error(`Internal XML error: ${parseError.textContent}`);
      }

      spTree.appendChild(doc.importNode(picDoc.documentElement, true));

      zip.file(
        "ppt/slides/slide1.xml",
        new XMLSerializer().serializeToString(doc),
      );

      const ctFile = zip.file("[Content_Types].xml");
      if (ctFile) {
        let ctXml = await ctFile.async("string");
        const ensureExt = (e: string, contentType: string) => {
          if (!ctXml.includes(`Extension="${e}"`)) {
            const entry = `<Default Extension="${e}" ContentType="${contentType}"/>`;
            ctXml = ctXml.replace(/(<Types[^>]*>)/, `$1${entry}`);
          }
        };

        if (isSvg) {
          ensureExt("png", "image/png");
          ensureExt("svg", "image/svg+xml");
        } else {
          ensureExt(ext, mime);
        }
        zip.file("[Content_Types].xml", ctXml);
      }

      markDirty();
      return { shapeId: String(shapeId), shapeName };
    });
  });
}

const insertImageCmd: DescribedCommand = {
  promptSnippet:
    "- insert-image <file> [<slide>] [--slide-id=ID] [--directory-version=VERSION] [--x=N] [--y=N] [--width=N] [--height=N] [--unit=pt|in|cm|emu] [--name=NAME] — Insert a VFS image using a stable slide ID (preferred) or legacy 1-based slide number; returns a JSON mutation receipt",
  command: {
    name: "insert-image",
    load: async () =>
      defineCommand("insert-image", async (args, ctx) => {
        const flags: Record<string, string> = {};
        const positional: string[] = [];
        for (const arg of args) {
          const match = arg.match(/^--([\w-]+)=(.+)$/);
          if (match) {
            flags[match[1]] = match[2];
          } else {
            positional.push(arg);
          }
        }

        try {
          if (positional.length < 1 || positional.length > 2) {
            throw new CommandInputError(
              "INVALID_ARGUMENTS",
              "Usage: insert-image <file> [<slide>] [--slide-id=ID] [--directory-version=VERSION] [--x=N] [--y=N] [--width=N] [--height=N] [--unit=pt|in|cm|emu] [--name=SHAPE_NAME]",
            );
          }
          const [filePath, slideArg] = positional;
          const { target } = parseCommandSlideTarget(slideArg, flags);
          const unit = flags.unit || "pt";
          if (!["in", "cm", "pt", "emu"].includes(unit)) {
            throw new CommandInputError(
              "INVALID_UNIT",
              "Unit must be one of: in, cm, pt, emu.",
            );
          }

          const x = flags.x ? Number.parseFloat(flags.x) : 0;
          const y = flags.y ? Number.parseFloat(flags.y) : 0;
          const width = flags.width
            ? Number.parseFloat(flags.width)
            : undefined;
          const height = flags.height
            ? Number.parseFloat(flags.height)
            : undefined;
          const widthProvided = width !== undefined;
          const heightProvided = height !== undefined;
          if (
            [x, y, width, height].some(
              (value) => value !== undefined && !Number.isFinite(value),
            ) ||
            (width !== undefined && width <= 0) ||
            (height !== undefined && height <= 0)
          ) {
            throw new CommandInputError(
              "INVALID_DIMENSIONS",
              "x and y must be finite numbers; width and height must be positive finite numbers.",
            );
          }

          const data = await resolveVfsPath(ctx, filePath);
          const { mime, ext } = detectMimeType(filePath, data);
          const fileName = filePath.split("/").pop() || "image";
          const shapeName = flags.name || fileName;
          const intrinsic = await getImageDimensions(data, ext);
          const resolvedSize = resolveImageSize({
            intrinsic,
            requestedWidth: width,
            requestedHeight: height,
            widthProvided,
            heightProvided,
            defaultWidth: pointsToUnit(DEFAULT_IMAGE_BOX_PT.width, unit),
            defaultHeight: pointsToUnit(DEFAULT_IMAGE_BOX_PT.height, unit),
          });

          const zipResult = await insertImageIntoSlide({
            target,
            data,
            mime,
            ext,
            shapeName,
            offX: parseUnit(x, unit),
            offY: parseUnit(y, unit),
            cx: parseUnit(resolvedSize.width, unit),
            cy: parseUnit(resolvedSize.height, unit),
          });

          const receipt = buildInsertReceipt("insert-image", zipResult, {
            source: filePath,
            mimeType: mime,
            x,
            y,
            width: resolvedSize.width,
            height: resolvedSize.height,
            unit,
          });
          return {
            stdout: JSON.stringify(receipt),
            stderr: "",
            exitCode: 0,
          };
        } catch (error) {
          return structuredCommandFailure(
            "insert-image",
            error,
            "INSERT_IMAGE_FAILED",
          );
        }
      }),
  },
};

const ICONIFY_API = "https://api.iconify.design";

const searchIconsCmd: DescribedCommand = {
  promptSnippet:
    "- search-icons <query> [--limit=N] [--prefix=ICON_SET] [--prefixes=SET1,SET2] — Search 200k+ vector icons from Iconify (Material, Fluent, Phosphor, Heroicons, Tabler, FontAwesome, etc.)",
  command: {
    name: "search-icons",
    load: async () =>
      defineCommand("search-icons", async (args) => {
        const flags: Record<string, string> = {};
        const positional: string[] = [];
        for (const arg of args) {
          const match = arg.match(/^--(\w+)=(.+)$/);
          if (match) {
            flags[match[1]] = match[2];
          } else {
            positional.push(arg);
          }
        }

        const query = positional.join(" ");
        if (!query) {
          return {
            stdout: "",
            stderr:
              "Usage: search-icons <query> [--limit=N] [--prefix=ICON_SET] [--prefixes=SET1,SET2]\n" +
              "  query      - Search term (e.g. 'warning', 'chart', 'home')\n" +
              "  --limit    - Max results (default: 32, min: 32, max: 999)\n" +
              "  --prefix   - Limit to one icon set (e.g. 'mdi', 'fluent')\n" +
              "  --prefixes - Comma-separated icon set prefixes (e.g. 'mdi,fluent,tabler')\n",
            exitCode: 1,
          };
        }

        try {
          const limit = flags.limit
            ? Math.max(32, Math.min(999, Number.parseInt(flags.limit, 10)))
            : 32;
          const params = new URLSearchParams({
            query,
            limit: String(limit),
          });
          if (flags.prefix) params.set("prefix", flags.prefix);
          if (flags.prefixes) params.set("prefixes", flags.prefixes);

          const res = await fetch(`${ICONIFY_API}/search?${params}`);
          if (!res.ok) {
            throw new Error(
              `Iconify API error: ${res.status} ${res.statusText}`,
            );
          }

          const data: {
            icons: string[];
            total: number;
            limit: number;
            start: number;
            collections: Record<
              string,
              {
                name: string;
                author?: { name: string };
                license?: { title: string };
                palette?: boolean;
              }
            >;
          } = await res.json();

          if (data.icons.length === 0) {
            return {
              stdout: "No icons found for that query.",
              stderr: "",
              exitCode: 0,
            };
          }

          const lines: string[] = [];
          lines.push(
            `Found ${data.total} icon(s) (showing ${data.icons.length}):\n`,
          );

          for (const iconId of data.icons) {
            const [prefix] = iconId.split(":");
            const col = data.collections[prefix];
            const setInfo = col ? ` [${col.name}]` : "";
            lines.push(`  ${iconId}${setInfo}`);
          }

          if (Object.keys(data.collections).length > 0) {
            lines.push("\nIcon sets in results:");
            for (const [prefix, info] of Object.entries(data.collections)) {
              const author = info.author ? ` by ${info.author.name}` : "";
              const license = info.license ? ` (${info.license.title})` : "";
              const palette = info.palette ? ", multi-color" : ", mono";
              lines.push(
                `  ${prefix}: ${info.name}${author}${license}${palette}`,
              );
            }
          }

          lines.push(
            "\nUse insert-icon to place an icon on a slide:",
            "  insert-icon <icon_id> --slide-id=<ID> [--directory-version=<VERSION>] [--x=N] [--y=N] [--width=N] [--height=N] [--color=#HEX]",
          );

          return { stdout: lines.join("\n"), stderr: "", exitCode: 0 };
        } catch (error) {
          const msg = error instanceof Error ? error.message : String(error);
          return { stdout: "", stderr: msg, exitCode: 1 };
        }
      }),
  },
};

const insertIconCmd: DescribedCommand = {
  promptSnippet:
    "- insert-icon <icon_id> [<slide>] [--slide-id=ID] [--directory-version=VERSION] [--x=N] [--y=N] [--width=N] [--height=N] [--unit=pt|in|cm|emu] [--color=#HEX] [--name=NAME] — Insert a vector icon using a stable slide ID (preferred) or legacy 1-based slide number; returns a JSON mutation receipt",
  command: {
    name: "insert-icon",
    load: async () =>
      defineCommand("insert-icon", async (args) => {
        const flags: Record<string, string> = {};
        const positional: string[] = [];
        for (const arg of args) {
          const match = arg.match(/^--([\w-]+)=(.+)$/);
          if (match) {
            flags[match[1]] = match[2];
          } else {
            positional.push(arg);
          }
        }

        try {
          if (positional.length < 1 || positional.length > 2) {
            throw new CommandInputError(
              "INVALID_ARGUMENTS",
              "Usage: insert-icon <icon_id> [<slide>] [--slide-id=ID] [--directory-version=VERSION] [--x=N] [--y=N] [--width=N] [--height=N] [--unit=pt|in|cm|emu] [--color=#HEX] [--name=SHAPE_NAME]",
            );
          }
          const [iconId, slideArg] = positional;
          const { target } = parseCommandSlideTarget(slideArg, flags);
          const parts = iconId.split(":");
          if (parts.length !== 2 || parts.some((part) => part.length === 0)) {
            throw new CommandInputError(
              "INVALID_ICON_ID",
              'Icon ID must be in "prefix:name" format (e.g. "mdi:alert"). Use search-icons to find icons.',
            );
          }

          const unit = flags.unit || "pt";
          if (!["in", "cm", "pt", "emu"].includes(unit)) {
            throw new CommandInputError(
              "INVALID_UNIT",
              "Unit must be one of: in, cm, pt, emu.",
            );
          }
          const x = flags.x ? Number.parseFloat(flags.x) : 0;
          const y = flags.y ? Number.parseFloat(flags.y) : 0;
          const width = flags.width
            ? Number.parseFloat(flags.width)
            : undefined;
          const height = flags.height
            ? Number.parseFloat(flags.height)
            : undefined;
          const widthProvided = width !== undefined;
          const heightProvided = height !== undefined;
          if (
            [x, y, width, height].some(
              (value) => value !== undefined && !Number.isFinite(value),
            ) ||
            (width !== undefined && width <= 0) ||
            (height !== undefined && height <= 0)
          ) {
            throw new CommandInputError(
              "INVALID_DIMENSIONS",
              "x and y must be finite numbers; width and height must be positive finite numbers.",
            );
          }

          const [prefix, name] = parts;
          const svgUrl = new URL(`${ICONIFY_API}/${prefix}/${name}.svg`);
          if (flags.color) {
            svgUrl.searchParams.set("color", flags.color);
          }

          const res = await fetch(svgUrl.toString());
          if (!res.ok) {
            if (res.status === 404) {
              throw new Error(
                `Icon "${iconId}" not found. Use search-icons to find valid icon IDs.`,
              );
            }
            throw new Error(
              `Failed to fetch icon: ${res.status} ${res.statusText}`,
            );
          }

          const svgText = await res.text();
          const svgData = new TextEncoder().encode(svgText);
          const shapeName = flags.name || iconId;
          const mediaCount = Date.now();
          const intrinsic = await getImageDimensions(svgData, "svg");
          const resolvedSize = resolveImageSize({
            intrinsic,
            requestedWidth: width,
            requestedHeight: height,
            widthProvided,
            heightProvided,
            defaultWidth: pointsToUnit(DEFAULT_ICON_BOX_PT.width, unit),
            defaultHeight: pointsToUnit(DEFAULT_ICON_BOX_PT.height, unit),
          });

          const zipResult = await insertImageIntoSlide({
            target,
            data: svgData,
            mime: "image/svg+xml",
            ext: "svg",
            shapeName,
            offX: parseUnit(x, unit),
            offY: parseUnit(y, unit),
            cx: parseUnit(resolvedSize.width, unit),
            cy: parseUnit(resolvedSize.height, unit),
            mediaPrefix: `icon_${mediaCount}`,
          });

          const receipt = buildInsertReceipt("insert-icon", zipResult, {
            iconId,
            color: flags.color ?? null,
            x,
            y,
            width: resolvedSize.width,
            height: resolvedSize.height,
            unit,
          });
          return {
            stdout: JSON.stringify(receipt),
            stderr: "",
            exitCode: 0,
          };
        } catch (error) {
          return structuredCommandFailure(
            "insert-icon",
            error,
            "INSERT_ICON_FAILED",
          );
        }
      }),
  },
};

export function getCustomCommands(ns: StorageNamespace): CustomCommandsResult {
  const local: DescribedCommand[] = [
    insertImageCmd,
    searchIconsCmd,
    insertIconCmd,
  ];
  const shared = getSharedCustomCommands({
    ns,
    includeImageSearch: true,
  });
  return {
    commands: [...shared.commands, ...local.map((d) => d.command)],
    promptSnippets: [
      ...shared.promptSnippets,
      ...local.map((d) => d.promptSnippet),
    ],
  };
}
