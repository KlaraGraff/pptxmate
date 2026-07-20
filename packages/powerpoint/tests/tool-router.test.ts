import { describe, expect, it } from "vitest";
import {
  needsPowerPointThemeContext,
  powerPointToolAllowlist,
  routePowerPointMessage,
} from "../src/lib/tool-router";

function routedToolNames(
  message: string,
  priorUserMessages: readonly string[] = [],
): string[] | null {
  const allow = powerPointToolAllowlist(message, { priorUserMessages });
  return allow ? Array.from(allow) : null;
}

describe("PowerPoint tool routing", () => {
  it("exposes only compact text tools for a normal text request", () => {
    const names = routedToolNames("翻译第 2 页标题")!;

    expect(names).toContain("read_slide_texts");
    expect(names).toContain("update_slide_text");
    expect(names).not.toContain("screenshot_slide");
    expect(names).not.toContain("execute_office_js");
    expect(names).not.toContain("edit_slide_master");
  });

  it("exposes text write tools for direct replacement wording", () => {
    const names = routedToolNames("把第 3 页的旧公司名改为新公司名")!;

    expect(routePowerPointMessage("把第 3 页的旧公司名改为新公司名")).toBe(
      "text",
    );
    expect(names).toContain("edit_slide_text");
    expect(names).toContain("update_slide_text");
    expect(names).not.toContain("execute_office_js");
  });

  it.each([
    "Replace ACME with Beta",
    "Change ACME to Beta",
  ])("keeps direct English replacement on compact text tools: %s", (message) => {
    expect(routePowerPointMessage(message)).toBe("text");
    expect(routedToolNames(message)).toContain("update_slide_text");
    expect(routedToolNames(message)).not.toContain("execute_office_js");
  });

  it("keeps direct formatting wording on the layout route", () => {
    expect(routePowerPointMessage("把第 3 页标题改为红色")).toBe("layout");
    expect(routedToolNames("把第 3 页标题改为红色")).toBeNull();
  });

  it("adds specialized content tools only for broad or container text work", () => {
    const names = routedToolNames("翻译整份 PPT，包括表格和图表")!;

    expect(names).toContain("execute_office_js");
    expect(names).toContain("edit_slide_xml");
    expect(names).toContain("edit_slide_chart");
    expect(names).not.toContain("edit_slide_master");
  });

  it("keeps ambiguous requests on discovery tools", () => {
    expect(routedToolNames("你可以帮我做什么？")).toEqual([
      "read",
      "bash",
      "list_slides",
      "read_slides",
      "read_slide_texts",
    ]);
  });

  it.each([
    "删除当前选中的 logo",
    "把表格第三行删掉",
    "让这两个框一样宽",
    "把这张图换成附件图片",
    "Remove the selected logo",
    "Delete the third row from the table",
    "Make these two boxes the same width",
    "Replace this image with the attached picture",
  ])("keeps an object mutation on the full tool route: %s", (message) => {
    expect(routePowerPointMessage(message)).toBe("layout");
    expect(routedToolNames(message)).toBeNull();
  });

  it("derives continuation routing from the active session only", () => {
    expect(
      routePowerPointMessage("继续下一页", {
        priorUserMessages: ["调整第 2 页布局"],
      }),
    ).toBe("layout");
    expect(
      routePowerPointMessage("继续下一页", {
        priorUserMessages: ["翻译第 2 页"],
      }),
    ).toBe("text");
    expect(routedToolNames("继续下一页", ["调整第 2 页布局"])).toBeNull();
    expect(routedToolNames("继续下一页", ["翻译第 2 页"])).toContain(
      "update_slide_text",
    );
  });

  it("loads theme metadata only for explicit global design work", () => {
    expect(needsPowerPointThemeContext("把第 3 页标题改成红色", "layout")).toBe(
      false,
    );
    expect(needsPowerPointThemeContext("修改整套 PPT 的主题色", "layout")).toBe(
      true,
    );
    expect(needsPowerPointThemeContext("Create a deck", "create")).toBe(true);
  });
});
