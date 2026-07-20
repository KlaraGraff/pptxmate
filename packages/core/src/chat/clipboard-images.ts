const MIME_EXTENSIONS: Record<string, string> = {
  "image/png": "png",
  "image/x-png": "png",
  "image/jpeg": "jpg",
  "image/jpg": "jpg",
  "image/gif": "gif",
  "image/webp": "webp",
  "image/bmp": "bmp",
};
const SUPPORTED_EXTENSIONS = new Set([
  "png",
  "jpg",
  "jpeg",
  "gif",
  "webp",
  "bmp",
]);
const EXTENSION_MIME_TYPES: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  webp: "image/webp",
  bmp: "image/bmp",
};

let fallbackId = 0;

function createPasteId(): string {
  const uuid = globalThis.crypto?.randomUUID?.();
  if (uuid) return uuid.slice(0, 12);
  fallbackId += 1;
  return `${Date.now()}-${fallbackId}`;
}

function imageExtension(file: File): string | undefined {
  const mime = file.type.toLowerCase();
  if (MIME_EXTENSIONS[mime]) return MIME_EXTENSIONS[mime];

  const extension = file.name.toLowerCase().split(".").pop() ?? "";
  if (SUPPORTED_EXTENSIONS.has(extension)) {
    return extension === "jpeg" ? "jpg" : extension;
  }
  return undefined;
}

function renameFile(file: File, name: string, extension: string): File {
  try {
    return new File([file], name, {
      type: file.type || EXTENSION_MIME_TYPES[extension] || "image/png",
      lastModified: file.lastModified,
    });
  } catch {
    try {
      Object.defineProperty(file, "name", { configurable: true, value: name });
    } catch {
      // Very old WebViews may expose an immutable File without a constructor.
      // Keep the original file instead of making paste fail completely.
    }
    return file;
  }
}

function isSupportedImage(file: File): boolean {
  const mime = file.type.toLowerCase();
  return (mime === "" || mime.startsWith("image/")) && !!imageExtension(file);
}

/**
 * Extract image files from a paste event and give every image a stable,
 * collision-resistant VFS name. The read tool identifies images by extension,
 * so unsupported clipboard formats are intentionally ignored here.
 */
export function extractClipboardImages(
  clipboard: DataTransfer | null | undefined,
  existingNames: readonly string[] = [],
  idFactory: () => string = createPasteId,
): File[] {
  if (!clipboard) return [];

  const candidates: File[] = [];
  const seen = new Set<File>();
  for (const item of Array.from(clipboard.items ?? [])) {
    if (item.kind !== "file") continue;
    const file = item.getAsFile();
    if (!file || seen.has(file) || !isSupportedImage(file)) continue;
    seen.add(file);
    candidates.push(file);
  }

  // Some WebViews expose clipboard files but leave clipboardData.items empty.
  if (candidates.length === 0) {
    for (const file of Array.from(clipboard.files ?? [])) {
      if (!seen.has(file) && isSupportedImage(file)) {
        seen.add(file);
        candidates.push(file);
      }
    }
  }

  const reserved = new Set(existingNames);
  return candidates.map((file) => {
    const extension = imageExtension(file);
    if (!extension) return file;

    const base = `pasted-image-${idFactory()}`;
    let name = `${base}.${extension}`;
    let suffix = 2;
    while (reserved.has(name)) {
      name = `${base}-${suffix}.${extension}`;
      suffix += 1;
    }
    reserved.add(name);
    return renameFile(file, name, extension);
  });
}

export function isImageFilename(name: string): boolean {
  const extension = name.toLowerCase().split(".").pop() ?? "";
  return SUPPORTED_EXTENSIONS.has(extension);
}
