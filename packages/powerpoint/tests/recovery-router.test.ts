import { describe, expect, it } from "vitest";
import {
  getPowerPointToolRecoveryInfo,
  normalizePowerPointToolArgsForReplay,
} from "../src/lib/recovery-router";

describe("PowerPoint recovery routing", () => {
  it("classifies slide image bash commands as stable layout writes", () => {
    expect(
      getPowerPointToolRecoveryInfo("bash", {
        command:
          "insert-image /home/user/uploads/a.png 8 --slide-id=slide-H --directory-version=directory-v1:abc",
      }),
    ).toMatchObject({
      effect: "write",
      mutationKind: "layout",
      scope: {
        slide_id: "slide-H",
        directory_version: "directory-v1:abc",
      },
    });
    expect(
      getPowerPointToolRecoveryInfo("bash", { command: "ls /home/user" }),
    ).toEqual({ effect: "read", verificationKinds: [] });
  });

  it("normalizes a replacement ID into the bash replay identity", () => {
    const derived = normalizePowerPointToolArgsForReplay("bash", {
      command:
        "insert-icon mdi:alert 8 --directory-version=old --slide-id=old-id",
      slide_id: "replacement-id",
    });
    expect(derived).toEqual({
      command: "insert-icon mdi:alert 8 --slide-id=replacement-id",
    });
    expect(
      normalizePowerPointToolArgsForReplay("bash", {
        command: "insert-icon mdi:alert 8 --slide-id=replacement-id",
      }),
    ).toEqual(derived);
  });

  it("uses only capability-matched automatic verifiers", () => {
    expect(getPowerPointToolRecoveryInfo("read_slide_text", {})).toMatchObject({
      effect: "read",
      verificationKinds: ["text"],
    });
    expect(getPowerPointToolRecoveryInfo("verify_slides", {})).toMatchObject({
      verificationKinds: ["layout"],
    });
    expect(getPowerPointToolRecoveryInfo("edit_slide_chart", {})).toMatchObject(
      {
        effect: "write",
        mutationKind: "arbitrary",
      },
    );
    expect(
      getPowerPointToolRecoveryInfo("execute_office_js", {}),
    ).toMatchObject({ effect: "write", mutationKind: "arbitrary" });
    expect(
      getPowerPointToolRecoveryInfo("edit_slide_master", {}),
    ).toMatchObject({
      mutationKind: "arbitrary",
    });
  });

  it("keeps safe batch shape scope without retaining text", () => {
    const info = getPowerPointToolRecoveryInfo("update_slide_text", {
      slide_index: 2,
      updates: [
        { shape_id: "7", text: "secret one" },
        { shape_id: "9", text: "secret two" },
      ],
    });

    expect(info.scope).toEqual({ slide_index: 2, shape_ids: "7,9" });
    expect(JSON.stringify(info)).not.toContain("secret");
  });

  it("uses a stable slide ID instead of a stale numeric position", () => {
    const info = getPowerPointToolRecoveryInfo("edit_slide_text", {
      slide_id: "slide-H",
      slide_index: 7,
      directoryVersion: "directory-v1:fnv1a32:11111111",
      shape_id: "9",
    });

    expect(info.scope).toEqual({
      slide_id: "slide-H",
      directory_version: "directory-v1:fnv1a32:11111111",
      shape_id: "9",
    });
  });

  it("recognizes post-write replacement IDs in recovery receipts", () => {
    const info = getPowerPointToolRecoveryInfo("edit_slide_xml", {
      _modifiedSlideId: "slide-H-replacement",
      _modifiedSlide: 6,
      originalSlideId: "slide-H",
      replacementSlideId: "slide-H-replacement",
      directoryVersion: "directory-v1:fnv1a32:22222222",
    });

    expect(info.scope).toEqual({
      slide_id: "slide-H-replacement",
      directory_version: "directory-v1:fnv1a32:22222222",
      original_slide_id: "slide-H",
      replacement_slide_id: "slide-H-replacement",
    });
  });

  it("normalizes guards and omitted modes out of replay identity", () => {
    const guarded = normalizePowerPointToolArgsForReplay("edit_slide_text", {
      slide_index: 0,
      shape_id: "7",
      text: "Bonjour",
      expected_text_hash: "fnv1a32:12345678",
      explanation: "translate",
    });
    expect(guarded).toEqual({
      slide_index: 0,
      shape_id: "7",
      text: "Bonjour",
      mode: "replace",
    });

    const batch = normalizePowerPointToolArgsForReplay("update_slide_text", {
      slide_index: 0,
      updates: [
        {
          shape_id: "7",
          text: "Bonjour",
          expected_text: "Hello",
        },
      ],
    });
    expect(batch).toEqual({
      slide_index: 0,
      updates: [{ shape_id: "7", text: "Bonjour", mode: "replace" }],
    });
  });

  it("keeps replay identity stable when an ID moves to a new directory position", () => {
    const beforeMove = normalizePowerPointToolArgsForReplay("edit_slide_text", {
      slide_id: "slide-H",
      slide_index: 7,
      directory_version: "directory-v1:fnv1a32:aaaaaaaa",
      shape_id: "9",
      text: "Bonjour",
    });
    const afterMove = normalizePowerPointToolArgsForReplay("edit_slide_text", {
      slideId: "slide-H",
      slideIndex: 6,
      directoryVersion: "directory-v1:fnv1a32:bbbbbbbb",
      shape_id: "9",
      text: "Bonjour",
    });

    expect(beforeMove).toEqual({
      slide_id: "slide-H",
      shape_id: "9",
      text: "Bonjour",
      mode: "replace",
    });
    expect(afterMove).toEqual(beforeMove);
  });
});
