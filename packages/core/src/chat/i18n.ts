import { writable } from "svelte/store";

export type Locale = "en" | "zh-CN";

const LANGUAGE_KEY = "pptxmate-language";

const messages = {
  en: {
    "language.section": "Language",
    "language.label": "Display language",
    "language.english": "English",
    "language.chinese": "Chinese (Simplified)",
    "settings.apiConfiguration": "API configuration",
    "settings.provider": "Provider",
    "settings.selectProvider": "Select provider...",
    "settings.customEndpoint": "Custom Endpoint",
    "settings.apiType": "API Type",
    "settings.baseUrl": "Base URL",
    "settings.baseUrlHint": "The API endpoint URL for your provider",
    "settings.modelId": "Model ID",
    "settings.model": "Model",
    "settings.selectModel": "Select model...",
    "settings.authentication": "Authentication",
    "settings.apiKey": "API Key",
    "settings.enterApiKey": "Enter your API key",
    "settings.corsProxy": "CORS Proxy",
    "settings.corsProxyHint": "Required for Anthropic and some providers",
    "settings.enableCorsProxy": "Enable CORS proxy",
    "settings.disableCorsProxy": "Disable CORS proxy",
    "settings.proxyUrl": "Proxy URL",
    "settings.proxyUrlHint": "Your proxy should accept ?url=encoded_url format",
    "settings.thinkingLevel": "Thinking Level",
    "settings.thinkingHint": "Extended thinking for supported models",
    "settings.expandToolCalls": "Expand Tool Calls",
    "settings.expandToolCallsHint":
      "Show tool call details expanded by default",
    "settings.expandToolCallsOn": "Expand tool calls by default",
    "settings.expandToolCallsOff": "Collapse tool calls by default",
    "settings.webTools": "Web tools",
    "settings.searchProvider": "Default Search Provider",
    "settings.searchProviderHint": "Used by web-search.",
    "settings.imageSearchProvider": "Default Image Search Provider",
    "settings.imageSearchProviderHint": "Used by image-search.",
    "settings.fetchProvider": "Default Fetch Provider",
    "settings.fetchProviderHint": "Used by web-fetch.",
    "settings.requiredForBrave": "Required for Brave search",
    "settings.requiredForSerper": "Required for Serper search",
    "settings.requiredForExa": "Required for Exa search/fetch",
    "settings.optional": "Optional",
    "settings.showAdvancedKeys": "Show advanced saved API keys",
    "settings.hideAdvancedKeys": "Hide advanced saved API keys",
    "settings.using": "Using",
    "settings.viaOAuth": "via OAuth",
    "settings.configurePrompt": "Fill in all fields above to get started",
    "settings.agentSkills": "Agent skills",
    "settings.removeSkill": "Remove skill",
    "settings.noSkills": "No skills installed",
    "settings.installing": "Installing...",
    "settings.addFolder": "Add Folder",
    "settings.addFile": "Add File",
    "settings.skillsHint":
      "Add a skill folder or a single SKILL.md file. Skills must have valid frontmatter with name and description.",
    "settings.about": "About",
    "settings.aboutText":
      "uses your own API key to connect to LLM providers. Your key is stored locally in the browser.",
    "settings.customEndpointHint":
      "Custom Endpoint: Point to any OpenAI-compatible API (Ollama, vLLM, LMStudio) or other supported API types.",
    "settings.corsAbout":
      "CORS Proxy: Requests route through your proxy to bypass browser CORS restrictions. Required for Claude OAuth and some providers.",
    "settings.login": "Login",
    "settings.submit": "Submit",
    "settings.logout": "Logout",
    "settings.tryAgain": "Try again",
    "settings.oauthAwaitingCodex":
      "Complete login in the opened tab. The page will redirect to localhost and fail. Copy the full URL from your browser's address bar and paste it below:",
    "settings.oauthAwaiting":
      "Authorize in the opened tab, then paste the code shown on the redirect page:",
    "settings.oauthRedirectPlaceholder": "Paste the full redirect URL here",
    "settings.oauthCodePlaceholder": "Paste code#state here",
    "settings.oauthCorsHint":
      "Requires CORS proxy to be enabled for token exchange.",
    "settings.oauthExchanging": "Exchanging authorization code...",
    "settings.oauthConnected": "Connected via OAuth",
    "nav.newChat": "New Chat",
    "nav.deleteCurrentSession": "Delete Current Session",
    "nav.chat": "Chat",
    "nav.files": "Files",
    "nav.settings": "Settings",
    "nav.followOn": "Follow mode: ON",
    "nav.followOff": "Follow mode: OFF",
    "nav.lightMode": "Light mode",
    "nav.darkMode": "Dark mode",
    "nav.clearMessages": "Clear messages",
    "nav.dropFiles": "Drop files here",
    "nav.inputTokens": "Input tokens",
    "nav.outputTokens": "Output tokens",
    "nav.cacheRead": "Cache read tokens",
    "nav.cacheWrite": "Cache write tokens",
    "nav.totalCost": "Total cost",
    "nav.contextUsage": "Context usage",
    "chat.newChat": "New Chat",
    "chat.noMessages": "No messages",
    "chat.empty": "Start a conversation to get started",
    "chat.thinking": "thinking...",
    "chat.messagePlaceholder": "Type a message...",
    "chat.configureApi": "Configure API key in settings",
    "chat.uploadFiles": "Upload files",
    "chat.removeUpload": "Remove from list",
    "chat.inspectImage": "Please inspect the attached image.",
    "chat.inspectImages": "Please inspect the attached images.",
    "chat.inspectFile": "Please inspect the attached file.",
    "files.count": "{count} file{suffix}",
    "files.refresh": "Refresh",
    "files.empty": "No files in virtual filesystem",
    "files.emptyHint": "Upload files or let the agent create them",
    "files.preview": "Preview",
    "files.download": "Download",
    "files.delete": "Delete",
  },
  "zh-CN": {
    "language.section": "语言",
    "language.label": "界面语言",
    "language.english": "English",
    "language.chinese": "中文（简体）",
    "settings.apiConfiguration": "API 配置",
    "settings.provider": "服务商",
    "settings.selectProvider": "选择服务商...",
    "settings.customEndpoint": "自定义端点",
    "settings.apiType": "API 类型",
    "settings.baseUrl": "基础 URL",
    "settings.baseUrlHint": "服务商的 API 端点地址",
    "settings.modelId": "模型 ID",
    "settings.model": "模型",
    "settings.selectModel": "选择模型...",
    "settings.authentication": "认证方式",
    "settings.apiKey": "API 密钥",
    "settings.enterApiKey": "输入 API 密钥",
    "settings.corsProxy": "CORS 代理",
    "settings.corsProxyHint": "Anthropic 和部分服务商需要此项",
    "settings.enableCorsProxy": "启用 CORS 代理",
    "settings.disableCorsProxy": "关闭 CORS 代理",
    "settings.proxyUrl": "代理 URL",
    "settings.proxyUrlHint": "代理应接受 ?url=encoded_url 格式",
    "settings.thinkingLevel": "思考级别",
    "settings.thinkingHint": "为支持的模型启用扩展思考",
    "settings.expandToolCalls": "展开工具调用",
    "settings.expandToolCallsHint": "默认展开工具调用详情",
    "settings.expandToolCallsOn": "默认展开工具调用",
    "settings.expandToolCallsOff": "默认收起工具调用",
    "settings.webTools": "网络工具",
    "settings.searchProvider": "默认网页搜索服务商",
    "settings.searchProviderHint": "用于 web-search。",
    "settings.imageSearchProvider": "默认图片搜索服务商",
    "settings.imageSearchProviderHint": "用于 image-search。",
    "settings.fetchProvider": "默认网页抓取服务商",
    "settings.fetchProviderHint": "用于 web-fetch。",
    "settings.requiredForBrave": "Brave 搜索需要此密钥",
    "settings.requiredForSerper": "Serper 搜索需要此密钥",
    "settings.requiredForExa": "Exa 搜索/抓取需要此密钥",
    "settings.optional": "可选",
    "settings.showAdvancedKeys": "显示已保存的高级 API 密钥",
    "settings.hideAdvancedKeys": "隐藏已保存的高级 API 密钥",
    "settings.using": "当前使用",
    "settings.viaOAuth": "通过 OAuth",
    "settings.configurePrompt": "填写以上所有字段后即可开始",
    "settings.agentSkills": "智能体技能",
    "settings.removeSkill": "移除技能",
    "settings.noSkills": "尚未安装技能",
    "settings.installing": "正在安装...",
    "settings.addFolder": "添加文件夹",
    "settings.addFile": "添加文件",
    "settings.skillsHint":
      "添加技能文件夹或单个 SKILL.md 文件。技能必须包含有效的 name 和 description frontmatter。",
    "settings.about": "关于",
    "settings.aboutText":
      "使用您自己的 API 密钥连接 LLM 服务商。密钥仅存储在浏览器本地。",
    "settings.customEndpointHint":
      "自定义端点：可指向任何兼容 OpenAI 的 API（Ollama、vLLM、LMStudio）或其他受支持的 API 类型。",
    "settings.corsAbout":
      "CORS 代理：请求会通过代理绕过浏览器 CORS 限制。Claude OAuth 和部分服务商需要此项。",
    "settings.login": "登录",
    "settings.submit": "提交",
    "settings.logout": "退出登录",
    "settings.tryAgain": "重试",
    "settings.oauthAwaitingCodex":
      "请在打开的标签页中完成登录。页面会跳转到 localhost 并显示失败，请从浏览器地址栏复制完整 URL 并粘贴到下方：",
    "settings.oauthAwaiting":
      "请在打开的标签页中授权，然后粘贴跳转页显示的代码：",
    "settings.oauthRedirectPlaceholder": "在此粘贴完整跳转 URL",
    "settings.oauthCodePlaceholder": "粘贴 code#state",
    "settings.oauthCorsHint": "交换令牌时需要启用 CORS 代理。",
    "settings.oauthExchanging": "正在交换授权码...",
    "settings.oauthConnected": "已通过 OAuth 连接",
    "nav.newChat": "新建对话",
    "nav.deleteCurrentSession": "删除当前对话",
    "nav.chat": "对话",
    "nav.files": "文件",
    "nav.settings": "设置",
    "nav.followOn": "跟随模式：开启",
    "nav.followOff": "跟随模式：关闭",
    "nav.lightMode": "浅色模式",
    "nav.darkMode": "深色模式",
    "nav.clearMessages": "清空消息",
    "nav.dropFiles": "将文件拖放到此处",
    "nav.inputTokens": "输入令牌数",
    "nav.outputTokens": "输出令牌数",
    "nav.cacheRead": "缓存读取令牌数",
    "nav.cacheWrite": "缓存写入令牌数",
    "nav.totalCost": "总费用",
    "nav.contextUsage": "上下文用量",
    "chat.newChat": "新建对话",
    "chat.noMessages": "暂无消息",
    "chat.empty": "开始对话以继续",
    "chat.thinking": "思考中...",
    "chat.messagePlaceholder": "输入消息...",
    "chat.configureApi": "请先在设置中配置 API 密钥",
    "chat.uploadFiles": "上传文件",
    "chat.removeUpload": "从列表中移除",
    "chat.inspectImage": "请检查附加的图片。",
    "chat.inspectImages": "请检查附加的图片。",
    "chat.inspectFile": "请检查附加的文件。",
    "files.count": "{count} 个文件",
    "files.refresh": "刷新",
    "files.empty": "虚拟文件系统中没有文件",
    "files.emptyHint": "上传文件，或让智能体创建文件",
    "files.preview": "预览",
    "files.download": "下载",
    "files.delete": "删除",
  },
} as const;

export type TranslationKey = keyof (typeof messages)["en"];

function getInitialLocale(): Locale {
  if (typeof window === "undefined") return "en";

  const saved = localStorage.getItem(LANGUAGE_KEY);
  if (saved === "en" || saved === "zh-CN") return saved;
  return navigator.language.toLowerCase().startsWith("zh") ? "zh-CN" : "en";
}

export const locale = writable<Locale>(getInitialLocale());

export function setLocale(nextLocale: Locale) {
  locale.set(nextLocale);
  if (typeof window !== "undefined") {
    localStorage.setItem(LANGUAGE_KEY, nextLocale);
    document.documentElement.lang = nextLocale;
  }
}

export function applyLocale(nextLocale: Locale) {
  if (typeof document !== "undefined")
    document.documentElement.lang = nextLocale;
}

export function t(locale: Locale, key: TranslationKey): string {
  return messages[locale][key];
}

export function tf(
  locale: Locale,
  key: TranslationKey,
  replacements: Record<string, string | number>,
): string {
  return t(locale, key).replace(/\{(\w+)\}/g, (_, name: string) =>
    String(replacements[name] ?? `{${name}}`),
  );
}
