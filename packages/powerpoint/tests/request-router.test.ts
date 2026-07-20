import { describe, expect, it, vi } from "vitest";

vi.mock("@office-agents/core", () => ({
  buildSkillsPromptSection: () => "",
}));

import {
  type CompactContextMessage,
  compactPowerPointContext,
  isPowerPointContinuationRequest,
  routePowerPointRequest,
} from "../src/lib/request-router";
import { buildPowerPointSystemPrompt } from "../src/lib/system-prompt";

function textContent(text: string) {
  return [{ type: "text", text }];
}

function toolCallContent(id: string, code: string) {
  return [
    {
      type: "toolCall",
      id,
      name: "edit_slide_xml",
      arguments: { code },
    },
  ];
}

function toolResultContent(text: string) {
  return textContent(text);
}

describe("PowerPoint request routing", () => {
  it.each([
    ["请读取第 2 页的文字", "text"],
    ["Translate the slide title to French", "text"],
    ["翻译成法语并保持原格式不变", "text"],
    ["翻译成法语并保持加粗不变", "text"],
    ["Replace the wording without changing the font or color", "text"],
    ["为第 3 页生成法语翻译", "text"],
    ["检查这段翻译是否准确", "text"],
    ["检查第 3 页的翻译是否准确", "text"],
    ["检查 PPT 中的错别字", "text"],
    ["为这个 PPT 生成法语翻译", "text"],
    ["请总结这份 PPT", "text"],
    ["概括第 3 页", "text"],
    ["提取所有幻灯片内容", "text"],
    ["润色并改写第 2 页", "text"],
    ["读取第 3 页", "text"],
    ["把第 3 页的旧公司名改为新公司名", "text"],
    ["将 A 公司改成 B 公司", "text"],
    ["把第 5 页的 Open PPT 替换为 OpenPPT", "text"],
    ["Replace ACME with Beta", "text"],
    ["Change ACME to Beta", "text"],
    ["Read slide 3", "text"],
    ["Summarize this deck", "text"],
    ["Extract all slide copy", "text"],
    ["Polish slide 2", "text"],
    ["修改文本框位置，并保留字体颜色", "layout"],
    ["保持颜色并调整位置", "layout"],
    ["请调整第 2 页的字体和位置", "layout"],
    ["把第 3 页标题改成红色", "layout"],
    ["把第 3 页标题改为红色", "layout"],
    ["把所有黑色替换为蓝色", "layout"],
    ["把第 3 页标题加粗", "layout"],
    ["把第 3 页标题改成斜体", "layout"],
    ["把第 3 页标题换成微软雅黑", "layout"],
    ["删除第 3 页", "layout"],
    ["复制当前幻灯片", "layout"],
    ["把第 3 页移到最后", "layout"],
    ["重排所有幻灯片", "layout"],
    ["调整幻灯片顺序", "layout"],
    ["Delete slide 3", "layout"],
    ["Delete the third slide", "layout"],
    ["Duplicate the current slide", "layout"],
    ["Move slide 2 after slide 5", "layout"],
    ["Reorder the slides", "layout"],
    ["Adjust the slide order", "layout"],
    ["删除当前选中的 logo", "layout"],
    ["把表格第三行删掉", "layout"],
    ["让这两个框一样宽", "layout"],
    ["把这张图换成附件图片", "layout"],
    ["Remove the selected logo", "layout"],
    ["Delete the third row from the table", "layout"],
    ["Make these two boxes the same width", "layout"],
    ["Replace this image with the attached picture", "layout"],
    ["检查所有幻灯片是否有重叠", "verify"],
    ["Create a new presentation", "create"],
    ["请告诉我 PowerPoint 可以做什么", "general"],
  ])("routes %s to %s", (message, expected) => {
    expect(routePowerPointRequest(message)).toBe(expected);
  });

  it("gives creation and verification intents precedence over text/layout words", () => {
    expect(routePowerPointRequest("创建一份包含文字的演示文稿")).toBe("create");
    expect(routePowerPointRequest("验证布局中的文字是否重叠")).toBe("verify");
  });

  it("does not confuse copying slide text with duplicating a slide", () => {
    expect(routePowerPointRequest("复制第 3 页文字")).toBe("text");
    expect(routePowerPointRequest("Copy the text on slide 3")).toBe("text");
  });

  it("inherits the previous route for short continuation requests", () => {
    expect(routePowerPointRequest("继续", "text")).toBe("text");
    expect(routePowerPointRequest("下一页也一样", "layout")).toBe("layout");
    expect(routePowerPointRequest("继续下一页", "text")).toBe("text");
    expect(routePowerPointRequest("继续处理下一页", "layout")).toBe("layout");
    expect(routePowerPointRequest("接着做", "text")).toBe("text");
    expect(routePowerPointRequest("Move on to the next slide", "text")).toBe(
      "text",
    );
    expect(routePowerPointRequest("Move on to the next slide", "layout")).toBe(
      "layout",
    );
    expect(routePowerPointRequest("continue with the rest", "layout")).toBe(
      "layout",
    );
    expect(isPowerPointContinuationRequest("接着处理下一张")).toBe(true);
    expect(isPowerPointContinuationRequest("Move on to the next slide")).toBe(
      true,
    );
  });

  it("keeps general requests on a compact discovery prompt", () => {
    const general = buildPowerPointSystemPrompt([], [], "general");
    const text = buildPowerPointSystemPrompt([], [], "text");
    const layout = buildPowerPointSystemPrompt([], [], "layout");

    expect(general).toContain("lightweight discovery path");
    expect(general).not.toContain("## Office.js API Reference");
    expect(general.length).toBeLessThan(6_000);
    expect(text).toContain("text-only path");
    expect(layout).toContain("## Office.js API Reference");
  });

  it("defines stable ID semantics for original and current slide numbers", () => {
    for (const route of ["general", "text", "layout"] as const) {
      const prompt = buildPowerPointSystemPrompt([], [], route);
      expect(prompt).toContain("`slide_id` is authoritative");
      expect(prompt).toContain("`directoryVersion`");
      expect(prompt).toContain('"Original slide N"');
      expect(prompt).toContain('"Current slide N"');
      expect(prompt).toContain("`_modifiedSlideId`");
      expect(prompt).toContain("Do not reread text, fonts, colors, geometry");
      expect(prompt).toContain("use the injected `targetSlide`");
    }

    const layoutPrompt = buildPowerPointSystemPrompt([], [], "layout");
    expect(layoutPrompt).toContain("const shape = targetSlide.shapes.addTable");
    expect(layoutPrompt).not.toContain(
      "const slide = context.presentation.slides.getItemAt",
    );
  });
});

describe("compactPowerPointContext", () => {
  it("keeps only the newest presentation metadata block without mutating input", () => {
    const oldMetadata =
      '<ppt_context>\n{"slideCount": 1}\n</ppt_context>\n\n旧请求';
    const newestMetadata =
      '<ppt_context>\n{"slideCount": 2}\n</ppt_context>\n\n新请求';
    const messages: CompactContextMessage[] = [
      { role: "user", content: oldMetadata },
      { role: "assistant", content: textContent("旧答复") },
      { role: "user", content: newestMetadata },
      { role: "assistant", content: textContent("新答复") },
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 100_000,
      recentMessageCount: 10,
    });

    expect(compacted[0].content).toBe("旧请求");
    expect(compacted[2].content).toBe(newestMetadata);
    expect(messages[0].content).toBe(oldMetadata);
    expect(messages[2].content).toBe(newestMetadata);
  });

  it("removes stale metadata from real array-based user content", () => {
    const messages: CompactContextMessage[] = [
      {
        role: "user",
        content: [
          {
            type: "text",
            text: '<ppt_context>\n{"slideCount":1}\n</ppt_context>\n\n旧请求',
          },
        ],
      },
      { role: "assistant", content: textContent("旧答复") },
      {
        role: "user",
        content: [
          {
            type: "text",
            text: '<ppt_context>\n{"slideCount":2}\n</ppt_context>\n\n新请求',
          },
        ],
      },
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 100_000,
      recentMessageCount: 10,
    });

    expect((compacted[0].content as Array<{ text: string }>)[0].text).toBe(
      "旧请求",
    );
    expect(JSON.stringify(compacted[2].content)).toContain("slideCount");
    expect(JSON.stringify(messages[0].content)).toContain("ppt_context");
  });

  it("compacts large historical tool results while retaining their envelope", () => {
    const largeResult = "R".repeat(5_000);
    const messages: CompactContextMessage[] = [
      { role: "user", content: "旧请求" },
      {
        role: "toolResult",
        toolCallId: "old-result",
        content: toolResultContent(largeResult),
      },
      { role: "user", content: "最近请求" },
      { role: "assistant", content: textContent("最近答复") },
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 100_000,
      recentMessageCount: 2,
    });
    const result = compacted[1];

    expect(result.role).toBe("toolResult");
    expect(JSON.stringify(result)).toContain("compacted");
    expect(JSON.stringify(result)).not.toContain(largeResult);
  });

  it("removes old tool-call/result pairs but leaves recent messages intact", () => {
    const oldCallId = "old-call";
    const oldCode = `const xml = \`${"x".repeat(20_000)}\`;`;
    const recentMessages: CompactContextMessage[] = [
      { role: "user", content: "最近请求" },
      { role: "assistant", content: textContent("最近答复") },
    ];
    const messages: CompactContextMessage[] = [
      { role: "user", content: "旧请求" },
      {
        role: "assistant",
        content: toolCallContent(oldCallId, oldCode),
      },
      {
        role: "toolResult",
        toolCallId: oldCallId,
        content: toolResultContent("R".repeat(5_000)),
      },
      ...recentMessages,
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 3_000,
      recentMessageCount: recentMessages.length,
    });

    expect(
      compacted.some(
        (message) =>
          message.role === "toolResult" && message.toolCallId === oldCallId,
      ),
    ).toBe(false);
    expect(JSON.stringify(compacted)).not.toContain(oldCallId);
    expect(JSON.stringify(compacted)).not.toContain(oldCode);
    expect(compacted.slice(-recentMessages.length)).toEqual(recentMessages);
  });

  it("drops oversized reasoning blocks when one current turn exceeds budget", () => {
    const messages: CompactContextMessage[] = [
      { role: "user", content: [{ type: "text", text: "当前请求" }] },
      {
        role: "assistant",
        content: [
          { type: "thinking", thinking: "T".repeat(20_000) },
          { type: "text", text: "可见答复" },
        ],
      },
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 3_000,
      recentMessageCount: 10,
    });

    expect(JSON.stringify(compacted)).not.toContain("T".repeat(1_000));
    expect(JSON.stringify(compacted)).toContain("可见答复");
    expect(JSON.stringify(compacted).length).toBeLessThan(3_000);
  });

  it("preserves paired progress receipts for multiple large tools in the current turn", () => {
    const readCallId = "read-page-1";
    const writeCallId = "write-page-1";
    const newestCallId = "list-page-2";
    const fullText = `FULL_TEXT_${"T".repeat(12_000)}`;
    const rawXml = `<p:sld>${"X".repeat(12_000)}</p:sld>`;
    const editCode = `const xml = \`${"C".repeat(12_000)}\`;`;
    const messages: CompactContextMessage[] = [
      { role: "user", content: "继续翻译剩余幻灯片" },
      {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: readCallId,
            name: "read_slides",
            arguments: { slide_ids: ["slide-1", "slide-2"], code: editCode },
          },
        ],
      },
      {
        role: "toolResult",
        toolCallId: readCallId,
        content: toolResultContent(
          JSON.stringify({
            success: true,
            result: {
              slideIndex: 1,
              items: [{ textPreview: fullText, rawOoxml: rawXml }],
              hasMore: true,
              remainingSlideIds: ["slide-3", "slide-4"],
            },
          }),
        ),
      },
      {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: writeCallId,
            name: "update_slide_text",
            arguments: { code: editCode, text: fullText },
          },
        ],
      },
      {
        role: "toolResult",
        toolCallId: writeCallId,
        content: toolResultContent(
          JSON.stringify({
            success: true,
            slideIndex: 1,
            updatedShapeIds: ["shape-8", "shape-9"],
            rawOoxml: rawXml,
          }),
        ),
      },
      {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: newestCallId,
            name: "list_slides",
            arguments: { offset: 25, limit: 25 },
          },
        ],
      },
      {
        role: "toolResult",
        toolCallId: newestCallId,
        content: toolResultContent(
          JSON.stringify({
            success: true,
            result: {
              page: { hasMore: false, nextOffset: null },
            },
          }),
        ),
      },
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 2_400,
      recentMessageCount: 20,
    });
    const serialized = JSON.stringify(compacted);
    const callIds = new Set(
      compacted.flatMap((message) =>
        message.role === "assistant" && Array.isArray(message.content)
          ? (message.content as Array<Record<string, unknown>>)
              .filter((block) => block.type === "toolCall")
              .map((block) => block.id)
          : [],
      ),
    );
    const resultIds = compacted
      .filter((message) => message.role === "toolResult")
      .map((message) => message.toolCallId);
    const receiptFor = (toolCallId: string) => {
      const message = compacted.find(
        (candidate) =>
          candidate.role === "toolResult" &&
          candidate.toolCallId === toolCallId,
      );
      const block = (message?.content as Array<{ text: string }>)[0];
      return JSON.parse(block.text) as Record<string, unknown>;
    };

    expect(serialized.length).toBeLessThanOrEqual(2_400);
    expect(resultIds).toEqual([readCallId, writeCallId, newestCallId]);
    expect(
      resultIds.every((id) => typeof id === "string" && callIds.has(id)),
    ).toBe(true);
    expect(receiptFor(readCallId)).toEqual({
      toolName: "read_slides",
      toolCallId: readCallId,
      success: true,
      scope: { slideIndex: 1 },
      hasMore: true,
      remainingSlideIds: ["slide-3", "slide-4"],
    });
    expect(receiptFor(writeCallId)).toEqual({
      toolName: "update_slide_text",
      toolCallId: writeCallId,
      success: true,
      scope: { slideIndex: 1 },
      updatedShapeIds: ["shape-8", "shape-9"],
    });
    expect(receiptFor(newestCallId)).toEqual({
      success: true,
      result: { page: { hasMore: false, nextOffset: null } },
    });
    expect(serialized).not.toContain("FULL_TEXT_");
    expect(serialized).not.toContain("p:sld");
    expect(serialized).not.toContain("const xml");
  });

  it("keeps safe read and range-edit cursors without retaining slide text", () => {
    const readCallId = "read-text-range";
    const editCallId = "edit-text-range";
    const newestCallId = "verify-range";
    const fullText = `SECRET_BODY_${"T".repeat(10_000)}`;
    const rawXml = `<a:p>${"X".repeat(10_000)}</a:p>`;
    const messages: CompactContextMessage[] = [
      { role: "user", content: "继续处理这个超长文本框" },
      {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: readCallId,
            name: "read_slide_text",
            arguments: { code: rawXml, expected_text: fullText },
          },
        ],
      },
      {
        role: "toolResult",
        toolCallId: readCallId,
        content: toolResultContent(
          JSON.stringify({
            success: true,
            result: {
              slideId: "slide-H",
              slideIndex: 6,
              positionOneIndexed: 7,
              directoryVersion: "directory-v1:fnv1a32:11111111",
              indexMismatch: true,
              shapeId: "shape-4",
              shapeTextHash: "fnv1a32:11111111",
              paragraphs: [
                {
                  text: fullText,
                  textHash: "fnv1a32:22222222",
                  editScope: {
                    paragraph_start: 3,
                    paragraph_end: 4,
                    char_start: 200,
                    char_end: 1_200,
                    code: rawXml,
                  },
                },
              ],
              page: {
                hasMore: true,
                nextOffset: 3,
                nextCharOffset: 1_200,
                nextCursor: {
                  paragraph_offset: 3,
                  char_offset: 1_200,
                  text: fullText,
                },
                textHash: fullText,
              },
              xml: rawXml,
            },
          }),
        ),
      },
      {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: editCallId,
            name: "edit_slide_text",
            arguments: { text: fullText, code: rawXml },
          },
        ],
      },
      {
        role: "toolResult",
        toolCallId: editCallId,
        content: toolResultContent(
          JSON.stringify({
            success: true,
            scope: {
              paragraph_start: 3,
              paragraph_end: 4,
              char_start: 200,
              char_end: 1_200,
              text: fullText,
            },
            beforeTextHash: "fnv1a32:22222222",
            afterTextHash: `sha256:${"a".repeat(64)}`,
            textHash: fullText,
            nextCursor: {
              paragraph_offset: 3,
              char_offset: 1_350,
              code: rawXml,
            },
            originalSlideId: "slide-H",
            replacementSlideId: "slide-H-replacement",
            directoryVersion: "directory-v1:fnv1a32:22222222",
            _modifiedSlideId: "slide-H-replacement",
            _modifiedSlide: 6,
            rawOoxml: rawXml,
          }),
        ),
      },
      {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: newestCallId,
            name: "read_slide_text",
            arguments: { paragraph_offset: 3, char_offset: 1_350 },
          },
        ],
      },
      {
        role: "toolResult",
        toolCallId: newestCallId,
        content: toolResultContent(
          JSON.stringify({ success: true, result: { hasMore: false } }),
        ),
      },
    ];

    const compacted = compactPowerPointContext(messages, {
      maxChars: 3_000,
      recentMessageCount: 20,
    });
    const serialized = JSON.stringify(compacted);
    const receiptFor = (toolCallId: string) => {
      const message = compacted.find(
        (candidate) =>
          candidate.role === "toolResult" &&
          candidate.toolCallId === toolCallId,
      );
      return JSON.parse(
        (message?.content as Array<{ text: string }>)[0].text,
      ) as Record<string, unknown>;
    };

    expect(serialized.length).toBeLessThanOrEqual(3_000);
    expect(receiptFor(readCallId)).toEqual({
      toolName: "read_slide_text",
      toolCallId: readCallId,
      success: true,
      scope: {
        slideIndex: 6,
        slideId: "slide-H",
        positionOneIndexed: 7,
        directoryVersion: "directory-v1:fnv1a32:11111111",
        indexMismatch: true,
        shapeId: "shape-4",
      },
      hasMore: true,
      nextOffset: 3,
      nextCharOffset: 1_200,
      nextCursor: { paragraph_offset: 3, char_offset: 1_200 },
      editScope: {
        paragraph_start: 3,
        paragraph_end: 4,
        char_start: 200,
        char_end: 1_200,
      },
      shapeTextHash: "fnv1a32:11111111",
      textHash: "fnv1a32:22222222",
    });
    expect(receiptFor(editCallId)).toEqual({
      toolName: "edit_slide_text",
      toolCallId: editCallId,
      success: true,
      scope: {
        paragraph_start: 3,
        paragraph_end: 4,
        char_start: 200,
        char_end: 1_200,
        slideId: "slide-H-replacement",
        originalSlideId: "slide-H",
        replacementSlideId: "slide-H-replacement",
        directoryVersion: "directory-v1:fnv1a32:22222222",
        slideIndex: 6,
      },
      nextCursor: { paragraph_offset: 3, char_offset: 1_350 },
      beforeTextHash: "fnv1a32:22222222",
      afterTextHash: `sha256:${"a".repeat(64)}`,
    });
    expect(serialized).not.toContain("SECRET_BODY_");
    expect(serialized).not.toContain("a:p");
    expect(serialized).not.toContain('"code"');
  });

  it("collapses older receipt envelopes before sacrificing the newest tool pair", () => {
    const messages: CompactContextMessage[] = [
      { role: "user", content: "继续处理" },
    ];
    for (let index = 0; index < 12; index++) {
      const toolCallId = `write-${index}`;
      messages.push(
        {
          role: "assistant",
          content: [
            {
              type: "toolCall",
              id: toolCallId,
              name: "update_slide_text",
              arguments: { code: `SECRET_CODE_${"C".repeat(2_000)}` },
            },
          ],
        },
        {
          role: "toolResult",
          toolCallId,
          content: toolResultContent(
            JSON.stringify({
              success: true,
              slideIndex: index,
              updatedShapeIds: [`shape-${index}`],
              rawOoxml: `SECRET_XML_${"X".repeat(2_000)}`,
            }),
          ),
        },
      );
    }

    const compacted = compactPowerPointContext(messages, {
      maxChars: 2_200,
      recentMessageCount: 30,
    });
    const serialized = JSON.stringify(compacted);
    const toolResults = compacted.filter(
      (message) => message.role === "toolResult",
    );
    const summaryBlock = compacted
      .flatMap((message) =>
        message.role === "assistant" && Array.isArray(message.content)
          ? (message.content as Array<{ type: string; text?: string }>)
          : [],
      )
      .find(
        (block) =>
          block.type === "text" && block.text?.includes("toolProgress"),
      );
    const summary = JSON.parse(summaryBlock?.text ?? "{}") as {
      toolProgress?: Array<Record<string, unknown>>;
    };

    expect(serialized.length).toBeLessThanOrEqual(2_200);
    expect(toolResults.map((message) => message.toolCallId)).toEqual([
      "write-11",
    ]);
    expect(serialized).toContain('"id":"write-11"');
    expect(summary.toolProgress).toHaveLength(11);
    expect(summary.toolProgress?.[0]).toEqual({
      toolName: "update_slide_text",
      toolCallId: "write-0",
      success: true,
      scope: { slideIndex: 0 },
      updatedShapeIds: ["shape-0"],
    });
    expect(serialized).not.toContain("SECRET_CODE_");
    expect(serialized).not.toContain("SECRET_XML_");
  });
});
