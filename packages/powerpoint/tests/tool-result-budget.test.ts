import { Type } from "@sinclair/typebox";
import { describe, expect, it, vi } from "vitest";

vi.mock("@office-agents/core", () => ({ resizeImage: vi.fn() }));

import {
  capText,
  isSerializedJsonWithinBudget,
  serializedJsonByteLength,
  TOOL_RESULT_MAX_BYTES,
  utf8ByteLength,
} from "../src/lib/tools/result-budget";
import { defineTool, toolError, toolSuccess } from "../src/lib/tools/types";

describe("PowerPoint tool result budgets", () => {
  it("keeps small text unchanged", () => {
    expect(capText("ok", 1_000)).toEqual({
      text: "ok",
      truncated: false,
      omittedBytes: 0,
    });
  });

  it("returns a UTF-8-safe preview for oversized text", () => {
    const capped = capText("中🙂".repeat(2_000), 1_000);
    expect(capped.truncated).toBe(true);
    expect(capped.text).not.toContain("�");
    expect(capped.omittedBytes).toBeGreaterThan(0);
    expect(
      new TextEncoder().encode(capped.text).byteLength,
    ).toBeLessThanOrEqual(1_000);
  });

  it("normalizes an unrealistically small budget", () => {
    const capped = capText("x".repeat(500), 1);
    expect(capped.text).toHaveLength(256);
    expect(capped.truncated).toBe(true);
  });

  it("measures the final JSON encoding, including escapes and UTF-8 bytes", () => {
    const value = { text: 'quote=" slash=\\ 中文🙂' };
    const serialized = JSON.stringify(value);

    expect(serializedJsonByteLength(value)).toBe(utf8ByteLength(serialized));
    expect(serializedJsonByteLength(value)).toBeGreaterThan(
      utf8ByteLength(value.text),
    );
    expect(
      isSerializedJsonWithinBudget(value, serializedJsonByteLength(value)),
    ).toBe(true);
    expect(
      isSerializedJsonWithinBudget(value, serializedJsonByteLength(value) - 1),
    ).toBe(false);
    expect(TOOL_RESULT_MAX_BYTES).toBe(24 * 1024);
  });

  it("caps oversized structured success results", () => {
    const result = toolSuccess({
      success: true,
      data: "R".repeat(50_000),
      slideId: "replacement-1",
      originalSlideId: "slide-1",
      replacementSlideId: "replacement-1",
      newSlideId: "duplicate-1",
      slideIndex: 3,
      directoryVersion: "directory-v1:test",
      mutationCompleted: true,
      mutationState: "completed",
    });
    const text =
      result.content[0].type === "text" ? result.content[0].text : "";
    const payload = JSON.parse(text);

    expect(new TextEncoder().encode(text).byteLength).toBeLessThanOrEqual(
      24 * 1024,
    );
    expect(payload).toMatchObject({
      truncated: true,
      slideId: "replacement-1",
      originalSlideId: "slide-1",
      replacementSlideId: "replacement-1",
      newSlideId: "duplicate-1",
      slideIndex: 3,
      directoryVersion: "directory-v1:test",
      mutationCompleted: true,
      mutationState: "completed",
    });
  });

  it("caps oversized tool errors without losing the error envelope", () => {
    const metadata = Object.assign(new Error("E".repeat(50_000)), {
      code: "SLIDE_MUTATION_UNCERTAIN",
      slideId: "slide-7",
      directoryVersion: "directory-v1:test",
      mutationCompleted: true,
      mutationState: "uncertain",
    });
    const result = toolError(metadata);
    const text =
      result.content[0].type === "text" ? result.content[0].text : "";
    const payload = JSON.parse(text);

    expect(new TextEncoder().encode(text).byteLength).toBeLessThanOrEqual(
      24 * 1024,
    );
    expect(payload).toMatchObject({ success: false, truncated: true });
    expect(payload).toMatchObject({
      code: "SLIDE_MUTATION_UNCERTAIN",
      slideId: "slide-7",
      directoryVersion: "directory-v1:test",
      mutationCompleted: true,
      mutationState: "uncertain",
    });
    expect(payload.errorPreview.length).toBeGreaterThan(0);
  });

  it("does not enter a PowerPoint tool after the run was aborted", async () => {
    const execute = vi.fn(async () => toolSuccess({ success: true }));
    const tool = defineTool({
      name: "test_tool",
      label: "Test Tool",
      description: "Test",
      parameters: Type.Object({}),
      execute,
    });
    const controller = new AbortController();
    controller.abort();

    await expect(
      tool.execute("call-1", {}, controller.signal),
    ).rejects.toMatchObject({ name: "AbortError" });
    expect(execute).not.toHaveBeenCalled();
  });
});
