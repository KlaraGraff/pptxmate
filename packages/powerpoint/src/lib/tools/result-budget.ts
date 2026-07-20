export interface CappedText {
  text: string;
  truncated: boolean;
  omittedBytes: number;
}

export const TOOL_RESULT_MAX_BYTES = 24 * 1024;

export function utf8ByteLength(text: string): number {
  return new TextEncoder().encode(text).byteLength;
}

export function serializedJsonByteLength(value: unknown): number {
  return utf8ByteLength(JSON.stringify(value) ?? "");
}

export function isSerializedJsonWithinBudget(
  value: unknown,
  maxBytes = TOOL_RESULT_MAX_BYTES,
): boolean {
  return serializedJsonByteLength(value) <= Math.max(0, Math.floor(maxBytes));
}

export function capText(text: string, maxBytes: number): CappedText {
  const limit = Math.max(256, Math.floor(maxBytes));
  const bytes = new TextEncoder().encode(text);
  if (bytes.byteLength <= limit) {
    return { text, truncated: false, omittedBytes: 0 };
  }
  let preview = "";
  for (let end = limit; end > 0; end--) {
    try {
      preview = new TextDecoder("utf-8", { fatal: true }).decode(
        bytes.slice(0, end),
      );
      break;
    } catch {
      // Back up to a complete UTF-8 code point.
    }
  }
  return {
    text: preview,
    truncated: true,
    omittedBytes:
      bytes.byteLength - new TextEncoder().encode(preview).byteLength,
  };
}
