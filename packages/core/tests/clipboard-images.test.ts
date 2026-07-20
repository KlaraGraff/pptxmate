import { describe, expect, it } from "vitest";
import {
  extractClipboardImages,
  isImageFilename,
} from "../src/chat/clipboard-images";

function imageFile(name: string, type: string): File {
  return new File([new Uint8Array([1, 2, 3])], name, { type });
}

function clipboard(
  items: Array<Partial<DataTransferItem>> = [],
  files: File[] = [],
): DataTransfer {
  return { items, files } as unknown as DataTransfer;
}

describe("clipboard image extraction", () => {
  it("ignores ordinary text and unsupported image formats", () => {
    const result = extractClipboardImages(
      clipboard([
        { kind: "string", type: "text/plain" },
        {
          kind: "file",
          type: "image/tiff",
          getAsFile: () => imageFile("a.tiff", "image/tiff"),
        },
      ]),
      [],
      () => "fixed",
    );

    expect(result).toEqual([]);
  });

  it("extracts multiple supported images and assigns unique names", () => {
    const png = imageFile("image.png", "image/png");
    const jpg = imageFile("image.jpg", "image/jpeg");
    const result = extractClipboardImages(
      clipboard([
        { kind: "file", type: png.type, getAsFile: () => png },
        { kind: "file", type: jpg.type, getAsFile: () => jpg },
      ]),
      [],
      () => "fixed",
    );

    expect(result.map((file) => file.name)).toEqual([
      "pasted-image-fixed.png",
      "pasted-image-fixed.jpg",
    ]);
    expect(result.map((file) => file.type)).toEqual([
      "image/png",
      "image/jpeg",
    ]);
  });

  it("falls back to clipboardData.files when items are unavailable", () => {
    const file = imageFile("image.webp", "image/webp");
    const result = extractClipboardImages(
      clipboard([], [file]),
      [],
      () => "fallback",
    );

    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("pasted-image-fallback.webp");
  });

  it("accepts an image file whose WebView omitted the MIME type", () => {
    const file = imageFile("clipboard.jpeg", "");
    const result = extractClipboardImages(
      clipboard([], [file]),
      [],
      () => "no-mime",
    );

    expect(result[0].name).toBe("pasted-image-no-mime.jpg");
    expect(result[0].type).toBe("image/jpeg");
  });

  it("avoids names already present in the upload list", () => {
    const result = extractClipboardImages(
      clipboard([
        {
          kind: "file",
          type: "image/bmp",
          getAsFile: () => imageFile("image.bmp", "image/bmp"),
        },
      ]),
      ["pasted-image-fixed.bmp"],
      () => "fixed",
    );

    expect(result[0].name).toBe("pasted-image-fixed-2.bmp");
  });

  it("recognizes image attachment names for the composer", () => {
    expect(isImageFilename("pasted-image-a.png")).toBe(true);
    expect(isImageFilename("notes.txt")).toBe(false);
  });
});
