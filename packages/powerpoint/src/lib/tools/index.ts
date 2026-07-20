import type { AgentContext } from "@office-agents/core";
import { createBashTool, createReadTool } from "@office-agents/core";
import { duplicateSlideTool } from "./duplicate-slide";
import { createEditSlideChartTool } from "./edit-slide-chart";
import { createEditSlideMasterTool } from "./edit-slide-master";
import { editSlideTextTool } from "./edit-slide-text";
import { createEditSlideXmlTool } from "./edit-slide-xml";
import { createExecuteOfficeJsTool } from "./execute-office-js";
import { listSlideShapesTool } from "./list-slide-shapes";
import { listSlidesTool } from "./list-slides";
import { readSlideTextTool } from "./read-slide-text";
import { readSlideTextsTool } from "./read-slide-texts";
import { readSlidesTool } from "./read-slides";
import { screenshotSlideTool } from "./screenshot-slide";
import { updateSlideTextTool } from "./update-slide-text";
import { verifySlidesTool } from "./verify-slides";

export function createPptTools(ctx: AgentContext) {
  return [
    // fs tools
    createReadTool(ctx),
    createBashTool(ctx),
    // PPT read tools
    listSlidesTool,
    readSlidesTool,
    screenshotSlideTool,
    listSlideShapesTool,
    readSlideTextsTool,
    readSlideTextTool,
    verifySlidesTool,
    // PPT write tools
    createExecuteOfficeJsTool(ctx),
    editSlideTextTool,
    updateSlideTextTool,
    createEditSlideXmlTool(ctx),
    createEditSlideChartTool(ctx),
    createEditSlideMasterTool(ctx),
    duplicateSlideTool,
  ];
}

export {
  createBashTool,
  createReadTool,
  createEditSlideChartTool,
  createEditSlideMasterTool,
  createEditSlideXmlTool,
  createExecuteOfficeJsTool,
  duplicateSlideTool,
  editSlideTextTool,
  listSlidesTool,
  listSlideShapesTool,
  readSlidesTool,
  readSlideTextTool,
  readSlideTextsTool,
  screenshotSlideTool,
  updateSlideTextTool,
  verifySlidesTool,
};

export {
  defineTool,
  type ToolResult,
  toolError,
  toolImage,
  toolSuccess,
  toolText,
} from "./types";
