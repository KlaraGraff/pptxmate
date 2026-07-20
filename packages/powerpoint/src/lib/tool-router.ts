import type { MessagePreparationInfo } from "@office-agents/core";
import {
  isPowerPointContinuationRequest,
  type PowerPointTaskRoute,
  routePowerPointRequest,
} from "./request-router";

const TEXT_TOOL_NAMES = new Set([
  "read",
  "bash",
  "list_slides",
  "read_slides",
  "list_slide_shapes",
  "read_slide_texts",
  "read_slide_text",
  "edit_slide_text",
  "update_slide_text",
]);

const DISCOVERY_TOOL_NAMES = new Set([
  "read",
  "bash",
  "list_slides",
  "read_slides",
  "read_slide_texts",
]);

const SPECIALIZED_TEXT_RE =
  /(表格|图表|组合|整份|整套|全部幻灯片|全篇|全文|table|chart|group|whole deck|entire deck|all slides|deck[- ]wide)/i;
const SPECIALIZED_TEXT_TOOL_NAMES = [
  "execute_office_js",
  "edit_slide_xml",
  "edit_slide_chart",
];
const THEME_CONTEXT_RE =
  /(母版|主题|模板|版式|全局样式|master|theme|template|slide layout)/i;

export function previousPowerPointRoute(
  info?: MessagePreparationInfo,
): PowerPointTaskRoute {
  const messages = info?.priorUserMessages ?? [];
  for (let index = messages.length - 1; index >= 0; index--) {
    const message = messages[index];
    if (isPowerPointContinuationRequest(message)) continue;
    return routePowerPointRequest(message);
  }
  return "general";
}

export function routePowerPointMessage(
  userMessage: string,
  info?: MessagePreparationInfo,
): PowerPointTaskRoute {
  return routePowerPointRequest(userMessage, previousPowerPointRoute(info));
}

/** Null means the full PowerPoint tool set is required for this route. */
export function powerPointToolAllowlist(
  userMessage: string,
  info?: MessagePreparationInfo,
): ReadonlySet<string> | null {
  const route = routePowerPointMessage(userMessage, info);
  if (route === "general") return DISCOVERY_TOOL_NAMES;
  if (route !== "text") return null;

  const allow = new Set(TEXT_TOOL_NAMES);
  if (SPECIALIZED_TEXT_RE.test(userMessage)) {
    for (const name of SPECIALIZED_TEXT_TOOL_NAMES) allow.add(name);
  }
  return allow;
}

export function needsPowerPointThemeContext(
  userMessage: string,
  route: PowerPointTaskRoute,
): boolean {
  return (
    route === "create" ||
    (route === "layout" && THEME_CONTEXT_RE.test(userMessage))
  );
}
