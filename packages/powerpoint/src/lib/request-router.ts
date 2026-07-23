/**
 * Coarse, deterministic routing for PowerPoint requests.
 *
 * This intentionally runs locally (rather than asking the model to classify
 * the request) so routing never consumes another model turn or token budget.
 */
export type PowerPointTaskRoute =
  | "text"
  | "translationAudit"
  | "layout"
  | "create"
  | "verify"
  | "general";

const CREATE_RE =
  /((创建|新建|从零).{0,16}(演示文稿|幻灯片|PPT)|(制作|生成|搭建).{0,8}(一份|一套|整套|整份|全新).{0,8}(演示文稿|幻灯片|PPT)|(create|build)\s+(a\s+|an\s+|new\s+|the\s+)?(presentation|slide deck|deck|slides?)|make\s+(a\s+|new\s+)(presentation|slide deck|deck)|new presentation)/i;
const VERIFY_RE =
  /(重叠|越界|溢出|错位|(检查|验证|校验|审查).{0,16}(布局|排版|重叠|越界|溢出|错位)|(verify|inspect|check|review).{0,24}(layout|overlap|overflow|out of bounds|misalign)|screenshot)/i;
const TRANSLATION_AUDIT_RE =
  /((检查|验证|校验|审查|审校|核查|review|audit|check|verify).{0,24}(翻译|译文|双语|translation)|(翻译|译文|双语|translation).{0,24}(遗漏|漏译|缺失|不匹配|准确|一致|完整|遗漏|missing|omission|mismatch|accuracy|consistent|complete|review|audit|check|verify))/i;
const STRUCTURE_RE =
  /((删除|删掉|移除|复制|拷贝|克隆|移动|交换)(第\s*\d+\s*(页|张)|这(一)?(页|张|幻灯片)|当前(页|幻灯片)|幻灯片)(?![^，。；！？,.!?]{0,8}(文字|文本|内容))|(第\s*\d+\s*(页|张)|这(一)?(页|张|幻灯片)|当前(页|幻灯片)|幻灯片)[^，。；！？,.!?]{0,12}(移到|移动到|挪到|放到|删除|删掉|移除)|(重排|重新排序).{0,8}(幻灯片|页面|PPT)|(调整|改变|修改).{0,8}(幻灯片|页面|PPT).{0,8}(顺序|次序)|(添加|插入).{0,8}(幻灯片|页面)|\b(delete|remove|duplicate|copy|clone|move)\s+(the\s+)?((current|first|second|third|fourth|fifth|last|\d+(st|nd|rd|th)?)\s+)?(slide|page)s?\b|\breorder\s+(the\s+)?(slides|pages)\b|\b(change|adjust|update)\s+(the\s+)?(slide|page)\s+(order|sequence)\b)/i;
const LAYOUT_RE =
  /(字体|字号|颜色|色彩|样式|格式|布局|位置|坐标|大小|尺寸|宽度|高度|等宽|等高|同宽|同高|一样宽|一样高|对齐|间距|主题|母版|背景|设计|排版|font|color|style|format|layout|position|size|width|height|align|spacing|theme|master|background|design)/i;
const FORMAT_RE =
  /(加粗|粗体|斜体|下划线|删除线|字重|(改为|改成|换成|替换为|设为|设置为|使用|采用|改用|换用).{0,12}((深|浅)?(红|蓝|绿|黄|黑|白|灰|紫|橙|粉|青|棕)色|微软雅黑|宋体|黑体|仿宋|楷体|等线|Arial|Calibri|Aptos|Times New Roman)|\b(bold|italic|underline|strikethrough)\b|\b(make|set|change|turn|recolor)\b.{0,24}\b(red|blue|green|yellow|black|white|gray|grey|purple|orange|pink|cyan|brown)\b)/i;
const OBJECT_MUTATION_RE =
  /((删除|删掉|移除|替换|换成|换掉|更换|添加|插入|移动|挪动|调整|修改|改变|设置|裁剪|旋转|组合|取消组合)[^，。；！？,.!?]{0,40}(logo|标志|徽标|图标|图片|照片|图像|表格|图表|形状|文本框|框|对象|元素|箭头|线条|行|列|单元格)|(logo|标志|徽标|图标|图片|照片|图像|表格|图表|形状|文本框|框|对象|元素|箭头|线条|行|列|单元格)[^，。；！？,.!?]{0,40}(删除|删掉|移除|替换|换成|换掉|更换|添加|插入|移动|挪动|调整|修改|改变|设置|裁剪|旋转|组合|取消组合)|\b(delete|remove|replace|swap|add|insert|move|reposition|resize|align|crop|rotate|group|ungroup|modify|change|set)\b[^,.;!?]{0,48}\b(logo|icon|image|picture|photo|table|chart|shape|text box|box|object|element|arrow|line|row|column|cell)s?\b|\b(logo|icon|image|picture|photo|table|chart|shape|text box|box|object|element|arrow|line|row|column|cell)s?\b[^,.;!?]{0,48}\b(delete|remove|replace|swap|add|insert|move|reposition|resize|align|crop|rotate|group|ungroup|modify|change|set)\b)/i;
const MUTATION_FALLBACK_RE =
  /(删除|删掉|移除|换掉|更换|添加|插入|挪动|裁剪|旋转|组合|取消组合|\b(delete|remove|replace|swap|add|insert|move|reposition|resize|align|crop|rotate|group|ungroup|modify|change|set)\b)/i;
const DIRECT_TEXT_RE =
  /((?:把|将)[^，。；！？,.!?]{1,80}(?:改为|改成|替换为)[^，。；！？,.!?]{1,80}|\b(?:replace\s+[^,.;!?]{1,80}\s+with|change\s+[^,.;!?]{1,80}\s+to)\s+[^,.;!?]{1,80})/i;
const TEXT_RE =
  /(文字|文本|标题|正文|段落|内容|翻译|译成|英译|中译|替换|错别字|拼写|语法|校对|审校|总结|概括|摘要|提取|抽取|梳理|润色|改写|重写|精简|扩写|阅读|读取|读出|分析|修改.*字|\b(?:read|summari[sz]e|summary|extract|rewrite|rephrase|polish|condense|proofread)\b|text|translate|translation|wording|copy|paragraph|title|body|spelling|grammar|typo)/i;
const PRESERVED_LAYOUT_CLAUSE_RE =
  /((保持|保留|维持|不改|不修改|不要改|无需改)[^，。；！？,.!?]{0,48}|(格式|样式|字体|字号|颜色|位置|布局|排版)[^，。；！？,.!?]{0,16}(不变|保持|保留|维持|不改|不修改)|(keep|preserve|retain|without changing|do not change)[^,.;!?]{0,64}|(format|style|font|color|position|layout)[^,.;!?]{0,24}(unchanged|intact|as is))/gi;
const CONTRASTED_LAYOUT_RE =
  /((但|并|同时|然后|以及).{0,8}(调整|修改|改变|设置|移动|放大|缩小|对齐|更换).{0,12}(字体|字号|颜色|样式|格式|布局|位置|坐标|大小|尺寸|对齐|间距|主题|母版|背景|设计|排版)|(but|and|then).{0,12}(adjust|change|set|move|resize|align|recolor).{0,16}(font|color|style|format|layout|position|size|spacing|theme|master|background|design))/i;
const SAME_AS_BEFORE_RE =
  /^(?:同上|按(?:照)?(?:上(?:一)?次|刚才|上面)(?:的方式)?)(?:处理|继续|做)?(?:第\s*\d+\s*(?:页|张)|(?:下(?:一)?|另(?:一|外)?|后续|剩余|其余)(?:页|张|页面|幻灯片)?)?[。.!！\s]*$/i;
const CONTINUATION_PATTERNS = [
  /^(?:继续(?:处理)?(?:下(?:一)?(?:页|张)|后续(?:页面|幻灯片)?|剩余(?:页面|幻灯片)?|其余(?:页面|幻灯片)?)?(?:也)?(?:一样|同样)?|接着(?:做|处理|继续)?(?:下(?:一)?(?:页|张))?|下(?:一)?(?:页|张)(?:(?:也)?(?:一样|同样)|继续(?:处理)?)?|按(?:照)?刚才(?:的方式)?(?:处理)?|同样处理|照此处理)[。.!！\s]*$/i,
  SAME_AS_BEFORE_RE,
  /^(?:continue(?:\s+(?:(?:with|to|on)\s+)?(?:the\s+)?(?:rest|remaining\s+(?:slides?|pages?)|next\s+(?:slide|page)))?|keep\s+going(?:\s+(?:(?:with|on)\s+)?(?:the\s+)?(?:rest|next\s+(?:slide|page)))?|carry\s+on|same\s+(?:for|on)\s+the\s+next\s+(?:slide|page)|do\s+(?:the\s+)?(?:rest|same\s+for\s+the\s+next\s+(?:slide|page)))[.!?\s]*$/i,
  /^(?:move\s+on\s+to\s+(?:the\s+)?next\s+(?:slide|page))[.!?\s]*$/i,
];

export function isPowerPointContinuationRequest(userMessage: string): boolean {
  const text = userMessage.trim();
  return CONTINUATION_PATTERNS.some((pattern) => pattern.test(text));
}

export function routePowerPointRequest(
  userMessage: string,
  previousRoute: PowerPointTaskRoute = "general",
): PowerPointTaskRoute {
  const text = userMessage.trim();
  if (CREATE_RE.test(text)) return "create";
  if (VERIFY_RE.test(text)) return "verify";
  if (TRANSLATION_AUDIT_RE.test(text)) return "translationAudit";
  if (STRUCTURE_RE.test(text)) return "layout";
  const positiveLayoutText = text.replace(PRESERVED_LAYOUT_CLAUSE_RE, "");
  if (
    LAYOUT_RE.test(positiveLayoutText) ||
    FORMAT_RE.test(positiveLayoutText) ||
    CONTRASTED_LAYOUT_RE.test(text) ||
    OBJECT_MUTATION_RE.test(text)
  ) {
    return "layout";
  }
  if (DIRECT_TEXT_RE.test(text) || TEXT_RE.test(text)) return "text";
  if (previousRoute !== "general" && isPowerPointContinuationRequest(text)) {
    return previousRoute;
  }
  if (MUTATION_FALLBACK_RE.test(text)) return "layout";
  return "general";
}

export interface CompactContextMessage {
  role: string;
  content?: unknown;
  toolCallId?: string;
}

const METADATA_RE = /<ppt_context>\r?\n[\s\S]*?\r?\n<\/ppt_context>\r?\n\r?\n/;
const RUNTIME_CONTROL_RE =
  /^(?:<runtime_continue>[\s\S]*<\/runtime_continue>|<runtime_recovery>[\s\S]*<\/runtime_recovery>)$/;
const COMPACT_RESULT_MARKER =
  "[Earlier PowerPoint tool output compacted to protect context.]";

function textLength(value: unknown): number {
  if (typeof value === "string") return value.length;
  try {
    return JSON.stringify(value).length;
  } catch {
    return 0;
  }
}

function contentToText(content: unknown): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return "";
  return content
    .filter(
      (block): block is { type: "text"; text: string } =>
        !!block &&
        typeof block === "object" &&
        block.type === "text" &&
        typeof block.text === "string",
    )
    .map((block) => block.text)
    .join("\n");
}

function mapTextContent(
  content: unknown,
  transform: (text: string) => string,
): unknown {
  if (typeof content === "string") return transform(content);
  if (!Array.isArray(content)) return content;
  return content.map((block) => {
    if (
      block &&
      typeof block === "object" &&
      block.type === "text" &&
      typeof block.text === "string"
    ) {
      return { ...block, text: transform(block.text) };
    }
    return block;
  });
}

function keepTrailingText(content: unknown, maxChars: number): unknown {
  if (typeof content === "string") {
    return maxChars > 0 ? content.slice(-maxChars) : "";
  }
  if (!Array.isArray(content)) return content;

  let remaining = maxChars;
  const result = [...content];
  for (let i = result.length - 1; i >= 0; i--) {
    const block = result[i];
    if (
      !block ||
      typeof block !== "object" ||
      block.type !== "text" ||
      typeof block.text !== "string"
    ) {
      continue;
    }
    const text = remaining > 0 ? block.text.slice(-remaining) : "";
    result[i] = { ...block, text };
    remaining = Math.max(0, remaining - text.length);
  }
  return result;
}

function isUserTurnBoundary(message: CompactContextMessage): boolean {
  return (
    message.role === "user" &&
    !RUNTIME_CONTROL_RE.test(contentToText(message.content).trim())
  );
}

type JsonRecord = Record<string, unknown>;

interface CompactToolCall {
  id: string;
  name: string;
}

const SAFE_SCOPE_FIELDS = new Map([
  ["_modifiedSlideId", "slideId"],
  ["modifiedSlideId", "slideId"],
  ["_modifiedSlide", "slideIndex"],
  ["modifiedSlide", "slideIndex"],
  ["slideIndex", "slideIndex"],
  ["slide_index", "slideIndex"],
  ["slideId", "slideId"],
  ["slide_id", "slideId"],
  ["currentSlideId", "slideId"],
  ["current_slide_id", "slideId"],
  ["currentSlideIndex", "slideIndex"],
  ["current_slide_index", "slideIndex"],
  ["positionOneIndexed", "positionOneIndexed"],
  ["position_one_indexed", "positionOneIndexed"],
  ["directoryVersion", "directoryVersion"],
  ["directory_version", "directoryVersion"],
  ["originalSlideId", "originalSlideId"],
  ["original_slide_id", "originalSlideId"],
  ["replacementSlideId", "replacementSlideId"],
  ["replacement_slide_id", "replacementSlideId"],
  ["sourceSlideId", "sourceSlideId"],
  ["source_slide_id", "sourceSlideId"],
  ["sourceSlideIndex", "sourceSlideIndex"],
  ["source_slide_index", "sourceSlideIndex"],
  ["newSlideId", "newSlideId"],
  ["new_slide_id", "newSlideId"],
  ["newSlideIndex", "newSlideIndex"],
  ["new_slide_index", "newSlideIndex"],
  ["indexMismatch", "indexMismatch"],
  ["index_mismatch", "indexMismatch"],
  ["directoryChanged", "directoryChanged"],
  ["directory_changed", "directoryChanged"],
  ["inputDirectoryChanged", "inputDirectoryChanged"],
  ["input_directory_changed", "inputDirectoryChanged"],
  ["relocated", "relocated"],
  ["usedLegacyIndex", "usedLegacyIndex"],
  ["used_legacy_index", "usedLegacyIndex"],
  ["shapeId", "shapeId"],
  ["shape_id", "shapeId"],
  ["slideIndices", "slideIndices"],
  ["slide_indices", "slideIndices"],
  ["slideIds", "slideIds"],
  ["slide_ids", "slideIds"],
  ["shapeIds", "shapeIds"],
  ["shape_ids", "shapeIds"],
  ["paragraph_start", "paragraph_start"],
  ["paragraph_end", "paragraph_end"],
  ["char_start", "char_start"],
  ["char_end", "char_end"],
  ["offset", "offset"],
  ["limit", "limit"],
  ["returned", "returned"],
  ["total", "total"],
]);

function isJsonRecord(value: unknown): value is JsonRecord {
  return !!value && typeof value === "object" && !Array.isArray(value);
}

function parseToolResultContent(content: unknown): JsonRecord | undefined {
  const candidates =
    typeof content === "string"
      ? [content]
      : Array.isArray(content)
        ? content
            .filter(
              (block): block is { type: "text"; text: string } =>
                isJsonRecord(block) &&
                block.type === "text" &&
                typeof block.text === "string",
            )
            .map((block) => block.text)
        : [];

  for (const candidate of candidates) {
    try {
      const parsed: unknown = JSON.parse(candidate);
      if (isJsonRecord(parsed)) return parsed;
    } catch {
      // A non-JSON result has no safely extractable progress fields.
    }
  }
  return undefined;
}

function compactSafeValue(value: unknown): unknown {
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    value === null
  ) {
    return value;
  }
  if (typeof value === "string") {
    return value.replace(/\s+/g, " ").slice(0, 120);
  }
  if (Array.isArray(value)) {
    return value
      .filter(
        (item): item is string | number | boolean =>
          typeof item === "string" ||
          typeof item === "number" ||
          typeof item === "boolean",
      )
      .slice(0, 50)
      .map((item) =>
        typeof item === "string"
          ? item.replace(/\s+/g, " ").slice(0, 120)
          : item,
      );
  }
  return undefined;
}

function firstSafeField(
  records: Array<JsonRecord | undefined>,
  field: string,
): unknown {
  for (const record of records) {
    if (!record || !(field in record)) continue;
    const value = compactSafeValue(record[field]);
    if (value !== undefined) return value;
  }
  return undefined;
}

function firstSafeOffset(
  records: Array<JsonRecord | undefined>,
  field: string,
): number | null | undefined {
  for (const record of records) {
    if (!record || !(field in record)) continue;
    const value = record[field];
    if (value === null) return null;
    if (
      typeof value === "number" &&
      Number.isSafeInteger(value) &&
      value >= 0
    ) {
      return value;
    }
  }
  return undefined;
}

const SAFE_TEXT_HASH_RE =
  /^(?:fnv1a32:[0-9a-f]{8}|sha256:[0-9a-f]{64}|sha256-[a-zA-Z0-9+/]{43}=?)$/i;

function firstSafeTextHash(
  records: Array<JsonRecord | undefined>,
  field: string,
): string | undefined {
  for (const record of records) {
    const value = record?.[field];
    if (typeof value === "string" && SAFE_TEXT_HASH_RE.test(value)) {
      return value;
    }
  }
  return undefined;
}

function compactTextCursor(value: unknown): JsonRecord | null | undefined {
  if (value === null) return null;
  if (!isJsonRecord(value)) return undefined;
  const paragraphOffset = value.paragraph_offset;
  if (
    typeof paragraphOffset !== "number" ||
    !Number.isSafeInteger(paragraphOffset) ||
    paragraphOffset < 0
  ) {
    return undefined;
  }
  const cursor: JsonRecord = { paragraph_offset: paragraphOffset };
  if (
    typeof value.char_offset === "number" &&
    Number.isSafeInteger(value.char_offset) &&
    value.char_offset >= 0
  ) {
    cursor.char_offset = value.char_offset;
  }
  return cursor;
}

function firstTextCursor(
  records: Array<JsonRecord | undefined>,
): JsonRecord | null | undefined {
  for (const record of records) {
    if (!record || !("nextCursor" in record)) continue;
    const cursor = compactTextCursor(record.nextCursor);
    if (cursor !== undefined) return cursor;
  }
  return undefined;
}

function compactEditScope(value: unknown): JsonRecord | null | undefined {
  if (value === null) return null;
  if (!isJsonRecord(value)) return undefined;
  const paragraphStart = value.paragraph_start;
  const paragraphEnd = value.paragraph_end;
  if (
    typeof paragraphStart !== "number" ||
    !Number.isSafeInteger(paragraphStart) ||
    paragraphStart < 0 ||
    typeof paragraphEnd !== "number" ||
    !Number.isSafeInteger(paragraphEnd) ||
    paragraphEnd < paragraphStart
  ) {
    return undefined;
  }
  const scope: JsonRecord = {
    paragraph_start: paragraphStart,
    paragraph_end: paragraphEnd,
  };
  const charStart = value.char_start;
  const charEnd = value.char_end;
  if (
    typeof charStart === "number" &&
    Number.isSafeInteger(charStart) &&
    charStart >= 0 &&
    typeof charEnd === "number" &&
    Number.isSafeInteger(charEnd) &&
    charEnd >= charStart
  ) {
    scope.char_start = charStart;
    scope.char_end = charEnd;
  }
  return scope;
}

function firstEditScope(
  records: Array<JsonRecord | undefined>,
): JsonRecord | null | undefined {
  for (const record of records) {
    if (!record || !("editScope" in record)) continue;
    const scope = compactEditScope(record.editScope);
    if (scope !== undefined) return scope;
  }
  return undefined;
}

function buildToolProgressReceipt(
  toolCall: CompactToolCall,
  content: unknown,
): string {
  const parsed = parseToolResultContent(content);
  const payload = isJsonRecord(parsed?.result) ? parsed.result : undefined;
  const rootPage = isJsonRecord(parsed?.page) ? parsed.page : undefined;
  const payloadPage = isJsonRecord(payload?.page) ? payload.page : undefined;
  const records = [parsed, payload, rootPage, payloadPage];
  const paragraphRecords = Array.isArray(payload?.paragraphs)
    ? payload.paragraphs.filter(isJsonRecord).slice(0, 1)
    : [];
  const receipt: JsonRecord = {
    toolName: toolCall.name.replace(/[^a-zA-Z0-9_.-]/g, "").slice(0, 80),
    toolCallId: toolCall.id,
  };

  const success = firstSafeField([parsed], "success");
  if (typeof success === "boolean") receipt.success = success;

  const scope: JsonRecord = {};
  const explicitScopes = [parsed?.scope, payload?.scope].filter(isJsonRecord);
  for (const [sourceField, outputField] of SAFE_SCOPE_FIELDS) {
    const scopeRecords = [
      ...explicitScopes,
      parsed,
      payload,
      rootPage,
      payloadPage,
    ];
    const value = [
      "paragraph_start",
      "paragraph_end",
      "char_start",
      "char_end",
    ].includes(outputField)
      ? (() => {
          const rangeValue = firstSafeOffset(scopeRecords, sourceField);
          return typeof rangeValue === "number" ? rangeValue : undefined;
        })()
      : firstSafeField(scopeRecords, sourceField);
    if (value !== undefined && scope[outputField] === undefined) {
      scope[outputField] = value;
    }
  }
  if (Object.keys(scope).length > 0) receipt.scope = scope;

  const hasMore = firstSafeField(records, "hasMore");
  if (typeof hasMore === "boolean") receipt.hasMore = hasMore;
  const nextOffset = firstSafeOffset(records, "nextOffset");
  if (nextOffset !== undefined) receipt.nextOffset = nextOffset;
  const nextCharOffset = firstSafeOffset(records, "nextCharOffset");
  if (nextCharOffset !== undefined) receipt.nextCharOffset = nextCharOffset;
  const nextCursor = firstTextCursor(records);
  if (nextCursor !== undefined) receipt.nextCursor = nextCursor;
  const editScope = firstEditScope([...records, ...paragraphRecords]);
  if (editScope !== undefined) receipt.editScope = editScope;
  for (const field of ["remainingSlideIds", "updatedShapeIds"] as const) {
    const value = firstSafeField(records, field);
    if (Array.isArray(value)) receipt[field] = value;
  }
  for (const field of [
    "shapeTextHash",
    "textHash",
    "beforeTextHash",
    "afterTextHash",
  ] as const) {
    const value = firstSafeTextHash([...records, ...paragraphRecords], field);
    if (value) receipt[field] = value;
  }

  return JSON.stringify(receipt);
}

function minimalToolCallBlock(block: JsonRecord): JsonRecord {
  return {
    type: "toolCall",
    id: block.id,
    name: block.name,
    arguments: {},
  };
}

/**
 * Keep the transcript useful without sending every old OOXML/tool payload to
 * the provider. The persisted agent state remains untouched; this is only the
 * view passed to the next model request.
 */
export function compactPowerPointContext<T extends CompactContextMessage>(
  messages: T[],
  options: { maxChars?: number; recentMessageCount?: number } = {},
): T[] {
  const maxChars = options.maxChars ?? 60_000;
  const recentMessageCount = options.recentMessageCount ?? 18;
  const result = messages.map((message) => ({ ...message })) as T[];

  // A document snapshot is per-turn state. Keep only the newest snapshot.
  let newestMetadataIndex = -1;
  for (let i = result.length - 1; i >= 0; i--) {
    const message = result[i];
    if (
      message.role === "user" &&
      METADATA_RE.test(contentToText(message.content))
    ) {
      newestMetadataIndex = i;
      break;
    }
  }
  for (let i = 0; i < result.length; i++) {
    const message = result[i];
    if (i !== newestMetadataIndex && message.role === "user") {
      message.content = mapTextContent(message.content, (text) =>
        text.replace(METADATA_RE, ""),
      );
    }
  }

  const isOld = (index: number) => index < result.length - recentMessageCount;
  let estimatedChars = result.reduce(
    (sum, message) => sum + textLength(message.content),
    0,
  );

  // First pass: cap old tool results. Keeping the tool-result envelope and
  // toolCallId maintains provider message pairing.
  for (let i = 0; i < result.length; i++) {
    const message = result[i];
    if (!isOld(i) || message.role !== "toolResult") continue;
    const content = Array.isArray(message.content) ? message.content : [];
    const textBlocks = content.filter(
      (block): block is { type: "text"; text: string } =>
        !!block && typeof block === "object" && block.type === "text",
    );
    const originalText = textBlocks.map((block) => block.text).join("\n");
    if (originalText.length <= 2_000) continue;
    const replacement = `${originalText.slice(0, 1_600)}\n\n${COMPACT_RESULT_MARKER}`;
    message.content = [{ type: "text", text: replacement }];
    estimatedChars -= Math.max(0, originalText.length - replacement.length);
  }

  // Second pass: compact the oldest tool payloads until a conservative budget
  // is met. We do not delete messages, so tool-call/result ordering stays valid.
  if (estimatedChars > maxChars) {
    for (let i = 0; i < result.length && estimatedChars > maxChars; i++) {
      const message = result[i];
      if (message.role !== "toolResult" || !isOld(i)) continue;
      const content = Array.isArray(message.content) ? message.content : [];
      const oldSize = textLength(content);
      if (oldSize <= 180) continue;
      message.content = [
        { type: "text", text: COMPACT_RESULT_MARKER },
      ] as unknown as T["content"];
      estimatedChars -= oldSize - COMPACT_RESULT_MARKER.length;
    }
  }

  // Tool call arguments (especially execute_office_js/edit_slide_xml code)
  // can be much larger than their result. If the budget is still exceeded,
  // remove old tool-call/result pairs from the provider view as a valid,
  // human-readable assistant summary. The persisted transcript is untouched.
  if (estimatedChars > maxChars) {
    const omittedToolIds = new Set<string>();
    for (let i = 0; i < result.length && estimatedChars > maxChars; i++) {
      const message = result[i];
      if (
        !isOld(i) ||
        message.role !== "assistant" ||
        !Array.isArray(message.content)
      ) {
        continue;
      }
      const content = message.content as Array<Record<string, unknown>>;
      const hasToolCall = content.some((block) => block.type === "toolCall");
      const kept = content.filter((block) => {
        if (block.type === "toolCall") {
          const id = typeof block.id === "string" ? block.id : undefined;
          if (id) omittedToolIds.add(id);
          return false;
        }
        if (
          hasToolCall &&
          (block.type === "thinking" || block.type === "redactedThinking")
        ) {
          return false;
        }
        return true;
      });
      if (kept.length === content.length) continue;
      const oldSize = textLength(message.content);
      if (kept.length === 0) {
        message.content = [
          { type: "text", text: COMPACT_RESULT_MARKER },
        ] as unknown;
      } else {
        message.content = kept as unknown;
      }
      estimatedChars -= Math.max(0, oldSize - textLength(message.content));
    }
    if (omittedToolIds.size > 0) {
      for (let i = result.length - 1; i >= 0; i--) {
        const message = result[i];
        if (
          message.role === "toolResult" &&
          typeof message.toolCallId === "string" &&
          omittedToolIds.has(message.toolCallId)
        ) {
          result.splice(i, 1);
        }
      }
      estimatedChars = result.reduce(
        (sum, message) => sum + textLength(message.content),
        0,
      );
    }
  }

  // Final safety valve for unusually large user text or recent code payloads.
  // Cut only at a user-message boundary so the provider never receives a
  // dangling tool result without the turn that requested it.
  if (estimatedChars > maxChars) {
    let cut = 0;
    const preferredCut = Math.max(0, result.length - recentMessageCount);
    for (let i = preferredCut; i < result.length; i++) {
      if (isUserTurnBoundary(result[i])) {
        cut = i;
        break;
      }
    }
    if (cut > 0) result.splice(0, cut);
    estimatedChars = result.reduce(
      (sum, message) => sum + textLength(message.content),
      0,
    );
  }

  // If one current turn is itself enormous, preserve tool progress without
  // retaining code, XML, or full read payloads. Older pairs become minimal
  // provider-valid calls plus structured receipts; the newest pair remains
  // intact unless it alone would exceed the budget.
  if (Math.max(estimatedChars, textLength(result)) > maxChars) {
    const toolCalls = new Map<string, CompactToolCall>();
    for (const message of result) {
      if (message.role !== "assistant" || !Array.isArray(message.content)) {
        continue;
      }
      for (const block of message.content as JsonRecord[]) {
        if (
          block.type === "toolCall" &&
          typeof block.id === "string" &&
          typeof block.name === "string"
        ) {
          toolCalls.set(block.id, { id: block.id, name: block.name });
        }
      }
    }

    let newestPairedToolCallId: string | undefined;
    for (let i = result.length - 1; i >= 0; i--) {
      const message = result[i];
      if (
        message.role === "toolResult" &&
        typeof message.toolCallId === "string" &&
        toolCalls.has(message.toolCallId)
      ) {
        newestPairedToolCallId = message.toolCallId;
        break;
      }
    }

    const compactPair = (toolCallId: string) => {
      const toolCall = toolCalls.get(toolCallId);
      if (!toolCall) return;
      for (const message of result) {
        if (message.role === "assistant" && Array.isArray(message.content)) {
          message.content = (message.content as JsonRecord[]).map((block) =>
            block.type === "toolCall" && block.id === toolCallId
              ? minimalToolCallBlock(block)
              : block,
          ) as unknown;
        } else if (
          message.role === "toolResult" &&
          message.toolCallId === toolCallId
        ) {
          message.content = [
            {
              type: "text",
              text: buildToolProgressReceipt(toolCall, message.content),
            },
          ] as unknown as T["content"];
        }
      }
    };

    for (const message of result) {
      if (
        message.role === "toolResult" &&
        typeof message.toolCallId === "string" &&
        message.toolCallId !== newestPairedToolCallId &&
        toolCalls.has(message.toolCallId)
      ) {
        compactPair(message.toolCallId);
      }
    }

    const userPreviewChars = Math.min(
      8_000,
      Math.max(1_000, Math.floor(maxChars * 0.6)),
    );
    const assistantPreviewChars = Math.min(
      2_000,
      Math.max(500, Math.floor(maxChars * 0.2)),
    );
    for (const message of result) {
      if (message.role === "user") {
        message.content = keepTrailingText(
          message.content,
          userPreviewChars,
        ) as T["content"];
      } else if (
        message.role === "assistant" &&
        Array.isArray(message.content)
      ) {
        message.content = (message.content as Array<Record<string, unknown>>)
          .filter((block) => {
            return (
              block.type !== "thinking" && block.type !== "redactedThinking"
            );
          })
          .map((block) =>
            block.type === "text" && typeof block.text === "string"
              ? { ...block, text: block.text.slice(0, assistantPreviewChars) }
              : block,
          ) as unknown;
      }
    }

    let finalChars = textLength(result);
    if (finalChars > maxChars && newestPairedToolCallId) {
      compactPair(newestPairedToolCallId);
      finalChars = textLength(result);
    }

    while (finalChars > maxChars) {
      const nextUserIndex = result.findIndex(
        (message, index) => index > 0 && isUserTurnBoundary(message),
      );
      if (nextUserIndex <= 0) break;
      result.splice(0, nextUserIndex);
      finalChars = textLength(result);
    }

    // Tool receipts are the most valuable recovery state. If their envelopes
    // fit but ordinary prose does not, shed that prose before trimming the
    // current request.
    if (finalChars > maxChars) {
      for (const message of result) {
        if (message.role !== "assistant" || !Array.isArray(message.content)) {
          continue;
        }
        message.content = (message.content as JsonRecord[]).map((block) =>
          block.type === "text" ? { ...block, text: "" } : block,
        ) as unknown;
      }
      finalChars = textLength(result);
    }

    // Lots of tools can make their otherwise-small message envelopes costly.
    // If necessary, collapse complete older pairs into one assistant receipt;
    // the newest tool call/result pair remains in native provider form.
    if (finalChars > maxChars && newestPairedToolCallId) {
      const earlierToolCallIds = new Set(
        [...toolCalls.keys()].filter((id) => id !== newestPairedToolCallId),
      );
      const progress = result.flatMap((message) => {
        if (
          message.role !== "toolResult" ||
          typeof message.toolCallId !== "string" ||
          !earlierToolCallIds.has(message.toolCallId)
        ) {
          return [];
        }
        const parsed = parseToolResultContent(message.content);
        return parsed ? [parsed] : [];
      });

      if (progress.length > 0) {
        for (const message of result) {
          if (message.role !== "assistant" || !Array.isArray(message.content)) {
            continue;
          }
          message.content = (message.content as JsonRecord[]).filter(
            (block) => {
              if (block.type === "text" && block.text === "") return false;
              return (
                block.type !== "toolCall" ||
                typeof block.id !== "string" ||
                !earlierToolCallIds.has(block.id)
              );
            },
          ) as unknown;
        }
        for (let i = result.length - 1; i >= 0; i--) {
          const message = result[i];
          if (
            (message.role === "toolResult" &&
              typeof message.toolCallId === "string" &&
              earlierToolCallIds.has(message.toolCallId)) ||
            (message.role === "assistant" &&
              Array.isArray(message.content) &&
              message.content.length === 0)
          ) {
            result.splice(i, 1);
          }
        }
        const newestCallMessage = result.find(
          (message) =>
            message.role === "assistant" &&
            Array.isArray(message.content) &&
            (message.content as JsonRecord[]).some(
              (block) =>
                block.type === "toolCall" &&
                block.id === newestPairedToolCallId,
            ),
        );
        if (newestCallMessage && Array.isArray(newestCallMessage.content)) {
          newestCallMessage.content = [
            {
              type: "text",
              text: JSON.stringify({ toolProgress: progress }),
            },
            ...(newestCallMessage.content as JsonRecord[]),
          ] as unknown;
        }
        finalChars = textLength(result);
      }
    }
    if (finalChars > maxChars) {
      const newestUser = [...result]
        .reverse()
        .find((message) => isUserTurnBoundary(message));
      if (newestUser) {
        const currentTextChars = contentToText(newestUser.content).length;
        const targetTextChars = Math.max(
          0,
          currentTextChars - (finalChars - maxChars) - 64,
        );
        newestUser.content = keepTrailingText(
          newestUser.content,
          targetTextChars,
        ) as T["content"];
        finalChars = textLength(result);
        if (finalChars > maxChars) {
          newestUser.content = keepTrailingText(
            newestUser.content,
            0,
          ) as T["content"];
        }
      }
    }
  }

  return result;
}
