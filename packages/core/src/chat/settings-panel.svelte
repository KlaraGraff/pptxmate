<script lang="ts">
  import {
    API_TYPES,
    buildAuthorizationUrl,
    exchangeOAuthCode,
    generatePKCE,
    listFetchProviders,
    listImageSearchProviders,
    listSearchProviders,
    loadOAuthCredentials,
    loadSavedConfig,
    loadWebConfig,
    OAUTH_PROVIDERS,
    removeOAuthCredentials,
    saveConfig,
    saveOAuthCredentials,
    saveWebConfig,
    THINKING_LEVELS,
    type OAuthFlowState,
    type ThinkingLevel,
  } from "@office-agents/sdk";
  import {
    Check,
    ChevronDown,
    ChevronUp,
    ExternalLink,
    Eye,
    EyeOff,
    FolderUp,
    Globe2,
    LogOut,
    Plus,
    Trash2,
  } from "lucide-svelte";
  import { getChatContext } from "./chat-runtime-context";
  import { locale, setLocale, t, type Locale } from "./i18n";

  const chat = getChatContext();
  const runtimeState = chat.state;
  const adapter = chat.adapter;
  const ns = chat.context.namespace;

  let folderInputRef = $state<HTMLInputElement | null>(null);
  let fileInputRef = $state<HTMLInputElement | null>(null);
  let installing = $state(false);

  const saved = loadSavedConfig(ns);
  let provider = $state(saved?.provider || "");
  let apiKey = $state(saved?.apiKey || "");
  let model = $state(saved?.model || "");
  let showKey = $state(false);
  let useProxy = $state(saved?.useProxy !== false);
  let proxyUrl = $state(saved?.proxyUrl || "");
  let thinking = $state<ThinkingLevel>(saved?.thinking || "none");
  let apiType = $state(saved?.apiType || "openai-completions");
  let customBaseUrl = $state(saved?.customBaseUrl || "");
  let authMethod = $state<"apikey" | "oauth">(saved?.authMethod || "apikey");

  const savedWeb = loadWebConfig(ns);
  let webSearchProvider = $state(savedWeb.searchProvider);
  let imageSearchProvider = $state(savedWeb.imageSearchProvider);
  let webFetchProvider = $state(savedWeb.fetchProvider);
  let braveApiKey = $state(savedWeb.apiKeys.brave || "");
  let serperApiKey = $state(savedWeb.apiKeys.serper || "");
  let exaApiKey = $state(savedWeb.apiKeys.exa || "");
  let showAdvancedWebKeys = $state(false);

  let oauthFlow = $state<OAuthFlowState>(
    saved?.authMethod === "oauth"
      ? loadOAuthCredentials(ns, saved.provider)
        ? { step: "connected" }
        : { step: "idle" }
      : { step: "idle" },
  );
  let oauthCodeInput = $state("");

  const followMode = $derived($runtimeState.providerConfig?.followMode ?? true);
  const expandToolCalls = $derived(
    $runtimeState.providerConfig?.expandToolCalls ?? false,
  );
  const isCustom = $derived(provider === "custom");
  const models = $derived(
    provider && !isCustom ? chat.getModelsForProvider(provider) : [],
  );
  const hasOAuth = $derived(provider in OAUTH_PROVIDERS);
  const searchProviders = listSearchProviders();
  const imageSearchProviders = listImageSearchProviders();
  const fetchProviders = listFetchProviders();
  const needsBraveKey = $derived(webSearchProvider === "brave");
  const needsSerperKey = $derived(
    webSearchProvider === "serper" ||
      (adapter.hasImageSearch && imageSearchProvider === "serper"),
  );
  const needsExaKey = $derived(
    webSearchProvider === "exa" || webFetchProvider === "exa",
  );
  const isConfigured = $derived($runtimeState.providerConfig !== null);
  const showApiKeyInput = $derived(!(hasOAuth && authMethod === "oauth"));

  const inputStyle =
    "border-radius: var(--chat-radius); font-family: var(--chat-font-mono)";

  function updateAndSync(
    updates: Partial<{
      provider: string;
      apiKey: string;
      model: string;
      useProxy: boolean;
      proxyUrl: string;
      thinking: ThinkingLevel;
      apiType: string;
      customBaseUrl: string;
      authMethod: "apikey" | "oauth";
    }>,
  ) {
    const nextProvider = updates.provider ?? provider;
    const nextApiKey = updates.apiKey ?? apiKey;
    const nextModel = updates.model ?? model;
    const nextUseProxy = updates.useProxy ?? useProxy;
    const nextProxyUrl = updates.proxyUrl ?? proxyUrl;
    const nextThinking = updates.thinking ?? thinking;
    const nextApiType = updates.apiType ?? apiType;
    const nextCustomBaseUrl = updates.customBaseUrl ?? customBaseUrl;
    const nextAuthMethod = updates.authMethod ?? authMethod;

    provider = nextProvider;
    apiKey = nextApiKey;
    model = nextModel;
    useProxy = nextUseProxy;
    proxyUrl = nextProxyUrl;
    thinking = nextThinking;
    apiType = nextApiType;
    customBaseUrl = nextCustomBaseUrl;
    authMethod = nextAuthMethod;

    const isValid =
      nextProvider === "custom"
        ? Boolean(
            nextProvider &&
              nextApiType &&
              nextCustomBaseUrl &&
              nextModel &&
              nextApiKey,
          )
        : Boolean(nextProvider && nextApiKey && nextModel);

    if (!isValid) return;

    const config = {
      provider: nextProvider,
      apiKey: nextApiKey,
      model: nextModel,
      useProxy: nextUseProxy,
      proxyUrl: nextProxyUrl,
      thinking: nextThinking,
      followMode,
      expandToolCalls,
      apiType: nextApiType,
      customBaseUrl: nextCustomBaseUrl,
      authMethod: nextAuthMethod,
    };

    saveConfig(ns, config);
    chat.setProviderConfig(config);
  }

  function updateWebSettings(
    updates: Partial<{
      searchProvider: string;
      imageSearchProvider: string;
      fetchProvider: string;
      braveApiKey: string;
      serperApiKey: string;
      exaApiKey: string;
    }>,
  ) {
    webSearchProvider = updates.searchProvider ?? webSearchProvider;
    imageSearchProvider =
      updates.imageSearchProvider ?? imageSearchProvider;
    webFetchProvider = updates.fetchProvider ?? webFetchProvider;
    braveApiKey = updates.braveApiKey ?? braveApiKey;
    serperApiKey = updates.serperApiKey ?? serperApiKey;
    exaApiKey = updates.exaApiKey ?? exaApiKey;

    saveWebConfig(ns, {
      searchProvider: webSearchProvider,
      imageSearchProvider,
      fetchProvider: webFetchProvider,
      apiKeys: {
        brave: braveApiKey,
        serper: serperApiKey,
        exa: exaApiKey,
      },
    });
  }

  function handleProviderChange(newProvider: string) {
    if (newProvider === "custom") {
      updateAndSync({ provider: newProvider, model: "", authMethod: "apikey" });
    } else {
      const providerModels = newProvider
        ? chat.getModelsForProvider(newProvider)
        : [];
      const keepOAuth =
        newProvider in OAUTH_PROVIDERS ? authMethod : "apikey";
      updateAndSync({
        provider: newProvider,
        model: providerModels[0]?.id || "",
        authMethod: keepOAuth,
      });
    }

    if (!(newProvider in OAUTH_PROVIDERS)) {
      oauthFlow = { step: "idle" };
    }
  }

  function handleAuthMethodChange(newMethod: "apikey" | "oauth") {
    if (newMethod === "oauth") {
      const credentials = loadOAuthCredentials(ns, provider);
      if (credentials) {
        oauthFlow = { step: "connected" };
        updateAndSync({ authMethod: "oauth", apiKey: credentials.access });
      } else {
        authMethod = "oauth";
        oauthFlow = { step: "idle" };
      }
      return;
    }

    oauthFlow = { step: "idle" };
    updateAndSync({ authMethod: "apikey", apiKey: "" });
  }

  async function startOAuthLogin() {
    try {
      const { verifier, challenge } = await generatePKCE();
      const { url, oauthState } = buildAuthorizationUrl(
        provider,
        challenge,
        verifier,
      );
      window.open(url, "_blank");
      oauthFlow = { step: "awaiting-code", verifier, oauthState };
    } catch (error) {
      oauthFlow = {
        step: "error",
        message: error instanceof Error ? error.message : "Failed to start OAuth",
      };
    }
  }

  async function submitOAuthCode() {
    if (oauthFlow.step !== "awaiting-code" || !oauthCodeInput.trim()) return;

    const pendingFlow = oauthFlow;
    oauthFlow = { step: "exchanging" };

    try {
      const credentials = await exchangeOAuthCode({
        provider,
        rawInput: oauthCodeInput.trim(),
        verifier: pendingFlow.verifier,
        expectedState: pendingFlow.oauthState,
        useProxy,
        proxyUrl,
      });
      saveOAuthCredentials(ns, provider, credentials);
      oauthFlow = { step: "connected" };
      oauthCodeInput = "";
      updateAndSync({ apiKey: credentials.access, authMethod: "oauth" });
    } catch (error) {
      oauthFlow = {
        step: "error",
        message: error instanceof Error ? error.message : "OAuth failed",
      };
    }
  }

  function logoutOAuth() {
    removeOAuthCredentials(ns, provider);
    oauthFlow = { step: "idle" };
    updateAndSync({ authMethod: "apikey", apiKey: "" });
  }

  async function handleFolderSelect(event: Event) {
    const target = event.currentTarget as HTMLInputElement;
    const files = target.files;
    if (!files || files.length === 0) return;

    installing = true;
    try {
      await chat.installSkill(Array.from(files));
    } finally {
      installing = false;
      if (folderInputRef) folderInputRef.value = "";
    }
  }

  async function handleFileSelect(event: Event) {
    const target = event.currentTarget as HTMLInputElement;
    const files = target.files;
    if (!files || files.length === 0) return;

    installing = true;
    try {
      await chat.installSkill(Array.from(files));
    } finally {
      installing = false;
      if (fileInputRef) fileInputRef.value = "";
    }
  }
</script>

{#snippet toggleSwitch(active: boolean, onclick: () => void, ariaLabel: string)}
  <button
    type="button"
    {onclick}
    aria-label={ariaLabel}
    class={`w-10 h-5 rounded-full transition-colors relative ${active ? "bg-(--chat-accent)" : "bg-(--chat-border)"}`}
  >
    <span
      class={`absolute top-0.5 w-4 h-4 rounded-full bg-white transition-transform ${active ? "left-5" : "left-0.5"}`}
    ></span>
  </button>
{/snippet}

{#snippet apiKeyField(label: string, value: string, onInput: (v: string) => void, placeholder: string, altBg?: boolean)}
  <label class="block">
    <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
      {label}
    </span>
    <input
      type="password"
      {value}
      oninput={(e) => onInput((e.currentTarget as HTMLInputElement).value)}
      {placeholder}
      class={`w-full text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) placeholder:text-(--chat-text-muted) focus:outline-none focus:border-(--chat-border-active) ${altBg ? "bg-(--chat-bg)" : "bg-(--chat-input-bg)"}`}
      style={inputStyle}
    />
  </label>
{/snippet}

<div class="flex-1 overflow-y-auto p-4 space-y-6" style="font-family: var(--chat-font-mono)">
  <div>
    <div class="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-4">
      {t($locale, "language.section")}
    </div>
    <label class="block">
      <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
        {t($locale, "language.label")}
      </span>
      <div class="relative">
        <Globe2
          size={14}
          class="absolute left-3 top-1/2 -translate-y-1/2 text-(--chat-text-muted) pointer-events-none"
        />
        <select
          value={$locale}
          onchange={(event) =>
            setLocale((event.currentTarget as HTMLSelectElement).value as Locale)}
          class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm pl-9 pr-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active)"
          style={inputStyle}
        >
          <option value="en">{t($locale, "language.english")}</option>
          <option value="zh-CN">{t($locale, "language.chinese")}</option>
        </select>
      </div>
    </label>
  </div>

  <div>
    <div class="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-4">
      {t($locale, "settings.apiConfiguration")}
    </div>

    <div class="space-y-4">
      <label class="block">
        <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
          {t($locale, "settings.provider")}
        </span>
        <select
          value={provider}
          onchange={(event) =>
            handleProviderChange((event.currentTarget as HTMLSelectElement).value)}
          class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active)"
          style={inputStyle}
        >
          <option value="">{t($locale, "settings.selectProvider")}</option>
          {#each chat.availableProviders as availableProvider (availableProvider)}
            <option value={availableProvider}>{availableProvider}</option>
          {/each}
          <option disabled>──────────</option>
          <option value="custom">{t($locale, "settings.customEndpoint")}</option>
        </select>
      </label>

      {#if isCustom}
        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.apiType")}
          </span>
          <select
            value={apiType}
            onchange={(event) =>
              updateAndSync({
                apiType: (event.currentTarget as HTMLSelectElement).value,
              })}
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active)"
            style={inputStyle}
          >
            {#each API_TYPES as type (type.id)}
              <option value={type.id}>{type.name}</option>
            {/each}
          </select>
          <p class="text-[10px] text-(--chat-text-muted) mt-1">
            {API_TYPES.find((type) => type.id === apiType)?.hint}
          </p>
        </label>

        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.baseUrl")}
          </span>
          <input
            type="text"
            bind:value={customBaseUrl}
            oninput={() => updateAndSync({ customBaseUrl })}
            placeholder="https://api.openai.com/v1"
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) placeholder:text-(--chat-text-muted) focus:outline-none focus:border-(--chat-border-active)"
            style={inputStyle}
          />
          <p class="text-[10px] text-(--chat-text-muted) mt-1">
            {t($locale, "settings.baseUrlHint")}
          </p>
        </label>

        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.modelId")}
          </span>
          <input
            type="text"
            bind:value={model}
            oninput={() => updateAndSync({ model })}
            placeholder="gpt-4o"
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) placeholder:text-(--chat-text-muted) focus:outline-none focus:border-(--chat-border-active)"
            style={inputStyle}
          />
        </label>
      {/if}

      {#if !isCustom && provider}
        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.model")}
          </span>
          <select
            value={model}
            onchange={(event) =>
              updateAndSync({ model: (event.currentTarget as HTMLSelectElement).value })}
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active) disabled:opacity-50 disabled:cursor-not-allowed"
            style={inputStyle}
          >
            <option value="">{t($locale, "settings.selectModel")}</option>
            {#each models as availableModel (availableModel.id)}
              <option value={availableModel.id}>{availableModel.name}</option>
            {/each}
          </select>
        </label>
      {/if}

      {#if hasOAuth}
        <div>
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.authentication")}
          </span>
          <div class="flex gap-1">
            <button
              type="button"
              onclick={() => handleAuthMethodChange("apikey")}
              class={`flex-1 py-1.5 text-xs border transition-colors ${authMethod === "apikey" ? "bg-(--chat-accent) border-(--chat-accent) text-white" : "bg-(--chat-input-bg) border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active)"}`}
              style="border-radius: var(--chat-radius)"
            >
              {t($locale, "settings.apiKey")}
            </button>
            <button
              type="button"
              onclick={() => handleAuthMethodChange("oauth")}
              class={`flex-1 py-1.5 text-xs border transition-colors ${authMethod === "oauth" ? "bg-(--chat-accent) border-(--chat-accent) text-white" : "bg-(--chat-input-bg) border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active)"}`}
              style="border-radius: var(--chat-radius)"
            >
              {OAUTH_PROVIDERS[provider]?.label ?? "OAuth"}
            </button>
          </div>
        </div>
      {/if}

      {#if hasOAuth && authMethod === "oauth"}
        <div class="space-y-2">
          {#if oauthFlow.step === "idle"}
            <button
              type="button"
              onclick={startOAuthLogin}
              class="w-full flex items-center justify-center gap-2 px-3 py-2.5 text-xs bg-(--chat-input-bg) border border-(--chat-border) text-(--chat-text-primary) hover:border-(--chat-accent) hover:text-(--chat-accent) transition-colors"
              style="border-radius: var(--chat-radius)"
            >
              <ExternalLink size={12} />
              {OAUTH_PROVIDERS[provider]?.buttonText ?? t($locale, "settings.login")}
            </button>
          {:else if oauthFlow.step === "awaiting-code"}
            <div class="space-y-2">
              <p class="text-[10px] text-(--chat-text-muted)">
                {provider === "openai-codex"
                  ? t($locale, "settings.oauthAwaitingCodex")
                  : t($locale, "settings.oauthAwaiting")}
              </p>
              <div class="flex gap-1">
                <input
                  type="text"
                  bind:value={oauthCodeInput}
                  placeholder={provider === "openai-codex"
                    ? t($locale, "settings.oauthRedirectPlaceholder")
                    : t($locale, "settings.oauthCodePlaceholder")}
                  class="flex-1 bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) placeholder:text-(--chat-text-muted) focus:outline-none focus:border-(--chat-border-active)"
                  style={inputStyle}
                  onkeydown={(event) => event.key === "Enter" && submitOAuthCode()}
                />
                <button
                  type="button"
                  onclick={submitOAuthCode}
                  disabled={!oauthCodeInput.trim()}
                  class="px-3 py-2 text-xs bg-(--chat-accent) text-white border border-(--chat-accent) hover:opacity-90 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                  style="border-radius: var(--chat-radius)"
                >
                  {t($locale, "settings.submit")}
                </button>
              </div>
              <p class="text-[10px] text-(--chat-text-muted)">
                {t($locale, "settings.oauthCorsHint")}
              </p>
            </div>
          {:else if oauthFlow.step === "exchanging"}
            <div
              class="px-3 py-2.5 text-xs text-(--chat-text-muted) bg-(--chat-input-bg) border border-(--chat-border)"
              style="border-radius: var(--chat-radius)"
            >
              {t($locale, "settings.oauthExchanging")}
            </div>
          {:else if oauthFlow.step === "connected"}
            <div
              class="flex items-center justify-between px-3 py-2.5 bg-(--chat-input-bg) border border-(--chat-border)"
              style="border-radius: var(--chat-radius)"
            >
              <div class="flex items-center gap-2 text-xs">
                <Check size={12} class="text-(--chat-success)" />
                <span class="text-(--chat-text-secondary)">
                  {t($locale, "settings.oauthConnected")}
                </span>
              </div>
              <button
                type="button"
                onclick={logoutOAuth}
                class="flex items-center gap-1 text-[10px] text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
              >
                <LogOut size={10} />
                {t($locale, "settings.logout")}
              </button>
            </div>
          {:else if oauthFlow.step === "error"}
            <div class="space-y-2">
              <div
                class="px-3 py-2 text-xs text-(--chat-error) bg-(--chat-input-bg) border border-(--chat-error)/30"
                style="border-radius: var(--chat-radius)"
              >
                {oauthFlow.message}
              </div>
              <button
                type="button"
                onclick={() => (oauthFlow = { step: "idle" })}
                class="text-[10px] text-(--chat-text-muted) hover:text-(--chat-text-secondary) transition-colors"
              >
                {t($locale, "settings.tryAgain")}
              </button>
            </div>
          {/if}
        </div>
      {/if}

      {#if showApiKeyInput}
        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.apiKey")}
          </span>
          <div class="relative">
            <input
              type={showKey ? "text" : "password"}
              bind:value={apiKey}
              oninput={() => updateAndSync({ apiKey })}
              placeholder={t($locale, "settings.enterApiKey")}
              class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 pr-10 border border-(--chat-border) placeholder:text-(--chat-text-muted) focus:outline-none focus:border-(--chat-border-active)"
              style={inputStyle}
            />
            <button
              type="button"
              onclick={() => (showKey = !showKey)}
              class="absolute right-2 top-1/2 -translate-y-1/2 text-(--chat-text-muted) hover:text-(--chat-text-secondary)"
            >
              {#if showKey}
                <EyeOff size={14} />
              {:else}
                <Eye size={14} />
              {/if}
            </button>
          </div>
        </label>
      {/if}

      <div class="flex items-center justify-between">
        <div>
          <span class="text-xs text-(--chat-text-secondary)">
            {t($locale, "settings.corsProxy")}
          </span>
          <p class="text-[10px] text-(--chat-text-muted) mt-0.5">
            {t($locale, "settings.corsProxyHint")}
          </p>
        </div>
        {@render toggleSwitch(
          useProxy,
          () => updateAndSync({ useProxy: !useProxy }),
          useProxy
            ? t($locale, "settings.disableCorsProxy")
            : t($locale, "settings.enableCorsProxy"),
        )}
      </div>

      {#if useProxy}
        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.proxyUrl")}
          </span>
          <input
            type="text"
            bind:value={proxyUrl}
            oninput={() => updateAndSync({ proxyUrl })}
            placeholder="https://your-proxy.com/proxy"
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) placeholder:text-(--chat-text-muted) focus:outline-none focus:border-(--chat-border-active)"
            style={inputStyle}
          />
          <p class="text-[10px] text-(--chat-text-muted) mt-1">
            {t($locale, "settings.proxyUrlHint")}
          </p>
        </label>
      {/if}

      <div>
        <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
          {t($locale, "settings.thinkingLevel")}
        </span>
        <div class="flex gap-1">
          {#each THINKING_LEVELS as level (level.value)}
            <button
              type="button"
              onclick={() => updateAndSync({ thinking: level.value })}
              class={`flex-1 py-1.5 text-xs border transition-colors ${thinking === level.value ? "bg-(--chat-accent) border-(--chat-accent) text-white" : "bg-(--chat-input-bg) border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active)"}`}
              style="border-radius: var(--chat-radius)"
            >
              {level.label}
            </button>
          {/each}
        </div>
        <p class="text-[10px] text-(--chat-text-muted) mt-1">
          {t($locale, "settings.thinkingHint")}
        </p>
      </div>

      <div class="flex items-center justify-between">
        <div>
          <span class="text-xs text-(--chat-text-secondary)">
            {t($locale, "settings.expandToolCalls")}
          </span>
          <p class="text-[10px] text-(--chat-text-muted) mt-0.5">
            {t($locale, "settings.expandToolCallsHint")}
          </p>
        </div>
        {@render toggleSwitch(
          expandToolCalls,
          () => chat.toggleExpandToolCalls(),
          expandToolCalls
            ? t($locale, "settings.expandToolCallsOff")
            : t($locale, "settings.expandToolCallsOn"),
        )}
      </div>

      <div class="border-t border-(--chat-border) pt-4 space-y-3">
        <div class="text-[10px] uppercase tracking-widest text-(--chat-text-muted)">
          {t($locale, "settings.webTools")}
        </div>

        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.searchProvider")}
          </span>
          <select
            value={webSearchProvider}
            onchange={(event) =>
              updateWebSettings({
                searchProvider: (event.currentTarget as HTMLSelectElement).value,
              })}
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active)"
            style={inputStyle}
          >
            {#each searchProviders as searchProvider (searchProvider.id)}
              <option value={searchProvider.id}>{searchProvider.label}</option>
            {/each}
          </select>
          <p class="text-[10px] text-(--chat-text-muted) mt-1">
            {t($locale, "settings.searchProviderHint")}
          </p>
        </label>

        {#if adapter.hasImageSearch}
          <label class="block">
            <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
              {t($locale, "settings.imageSearchProvider")}
            </span>
            <select
              value={imageSearchProvider}
              onchange={(event) =>
                updateWebSettings({
                  imageSearchProvider:
                    (event.currentTarget as HTMLSelectElement).value,
                })}
              class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active)"
              style={inputStyle}
            >
              {#each imageSearchProviders as imageProvider (imageProvider.id)}
                <option value={imageProvider.id}>{imageProvider.label}</option>
              {/each}
            </select>
            <p class="text-[10px] text-(--chat-text-muted) mt-1">
              {t($locale, "settings.imageSearchProviderHint")}
            </p>
          </label>
        {/if}

        <label class="block">
          <span class="block text-xs text-(--chat-text-secondary) mb-1.5">
            {t($locale, "settings.fetchProvider")}
          </span>
          <select
            value={webFetchProvider}
            onchange={(event) =>
              updateWebSettings({
                fetchProvider: (event.currentTarget as HTMLSelectElement).value,
              })}
            class="w-full bg-(--chat-input-bg) text-(--chat-text-primary) text-sm px-3 py-2 border border-(--chat-border) focus:outline-none focus:border-(--chat-border-active)"
            style={inputStyle}
          >
            {#each fetchProviders as fetchProvider (fetchProvider)}
              <option value={fetchProvider}>{fetchProvider}</option>
            {/each}
          </select>
          <p class="text-[10px] text-(--chat-text-muted) mt-1">
            {t($locale, "settings.fetchProviderHint")}
          </p>
        </label>

        {#if needsBraveKey}
          {@render apiKeyField("Brave API Key", braveApiKey, (v) => { braveApiKey = v; updateWebSettings({ braveApiKey }); }, t($locale, "settings.requiredForBrave"))}
        {/if}

        {#if needsSerperKey}
          {@render apiKeyField("Serper API Key", serperApiKey, (v) => { serperApiKey = v; updateWebSettings({ serperApiKey }); }, t($locale, "settings.requiredForSerper"))}
        {/if}

        {#if needsExaKey}
          {@render apiKeyField("Exa API Key", exaApiKey, (v) => { exaApiKey = v; updateWebSettings({ exaApiKey }); }, t($locale, "settings.requiredForExa"))}
        {/if}

        <div class="pt-1">
          <button
            type="button"
            onclick={() => (showAdvancedWebKeys = !showAdvancedWebKeys)}
            class="inline-flex items-center gap-1.5 text-xs text-(--chat-text-secondary) hover:text-(--chat-text-primary)"
          >
            {#if showAdvancedWebKeys}
              <ChevronUp size={12} />
            {:else}
              <ChevronDown size={12} />
            {/if}
            <span>
              {showAdvancedWebKeys
                ? t($locale, "settings.hideAdvancedKeys")
                : t($locale, "settings.showAdvancedKeys")}
            </span>
          </button>
        </div>

        {#if showAdvancedWebKeys}
          <div class="space-y-3 border border-(--chat-border) p-3 bg-(--chat-input-bg)">
            {#if !needsBraveKey}
              {@render apiKeyField("Brave API Key", braveApiKey, (v) => { braveApiKey = v; updateWebSettings({ braveApiKey }); }, t($locale, "settings.optional"), true)}
            {/if}

            {#if !needsSerperKey}
              {@render apiKeyField("Serper API Key", serperApiKey, (v) => { serperApiKey = v; updateWebSettings({ serperApiKey }); }, t($locale, "settings.optional"), true)}
            {/if}

            {#if !needsExaKey}
              {@render apiKeyField("Exa API Key", exaApiKey, (v) => { exaApiKey = v; updateWebSettings({ exaApiKey }); }, t($locale, "settings.optional"), true)}
            {/if}
          </div>
        {/if}
      </div>
    </div>
  </div>

  <div class="border-t border-(--chat-border) pt-4">
    <div class="flex items-center gap-2 text-xs">
      {#if isConfigured}
        <Check size={12} class="text-(--chat-success)" />
        <span class="text-(--chat-text-secondary)">
          {t($locale, "settings.using")}
          {#if $runtimeState.providerConfig?.provider === "custom"}
            custom ({$runtimeState.providerConfig?.apiType})
          {:else}
            {$runtimeState.providerConfig?.provider}
          {/if}
          {$runtimeState.providerConfig?.authMethod === "oauth"
            ? ` ${t($locale, "settings.viaOAuth")}`
            : ""}
        </span>
      {:else}
        <span class="text-(--chat-text-muted)">
          {t($locale, "settings.configurePrompt")}
        </span>
      {/if}
    </div>
  </div>

  <div class="border-t border-(--chat-border) pt-4">
    <div class="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-4">
      {t($locale, "settings.agentSkills")}
    </div>

    <div class="space-y-3">
      {#if $runtimeState.skills.length > 0}
        <div class="space-y-1">
          {#each $runtimeState.skills as skill (skill.name)}
            <div
              class="flex items-start justify-between gap-2 px-3 py-2 bg-(--chat-input-bg) border border-(--chat-border)"
              style="border-radius: var(--chat-radius)"
            >
              <div class="min-w-0 flex-1">
                <div class="text-xs text-(--chat-text-primary) font-medium truncate">
                  {skill.name}
                </div>
                <div class="text-[10px] text-(--chat-text-muted) mt-0.5 line-clamp-2">
                  {skill.description}
                </div>
              </div>
              <button
                type="button"
                onclick={() => chat.uninstallSkill(skill.name)}
                class="shrink-0 p-1 text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
                title={t($locale, "settings.removeSkill")}
              >
                <Trash2 size={12} />
              </button>
            </div>
          {/each}
        </div>
      {:else}
        <p class="text-xs text-(--chat-text-muted)">
          {t($locale, "settings.noSkills")}
        </p>
      {/if}

      <div class="flex gap-2">
        <button
          type="button"
          onclick={() => folderInputRef?.click()}
          disabled={installing}
          class="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 text-xs bg-(--chat-input-bg) border border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active) hover:text-(--chat-text-primary) disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          style="border-radius: var(--chat-radius)"
        >
          <FolderUp size={12} />
          {installing
            ? t($locale, "settings.installing")
            : t($locale, "settings.addFolder")}
        </button>
        <button
          type="button"
          onclick={() => fileInputRef?.click()}
          disabled={installing}
          class="flex-1 flex items-center justify-center gap-1.5 px-3 py-2 text-xs bg-(--chat-input-bg) border border-(--chat-border) text-(--chat-text-secondary) hover:border-(--chat-border-active) hover:text-(--chat-text-primary) disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          style="border-radius: var(--chat-radius)"
        >
          <Plus size={12} />
          {installing
            ? t($locale, "settings.installing")
            : t($locale, "settings.addFile")}
        </button>
      </div>

      <p class="text-[10px] text-(--chat-text-muted)">
        {t($locale, "settings.skillsHint")}
      </p>
    </div>

    <input
      bind:this={folderInputRef}
      type="file"
      class="hidden"
      webkitdirectory={true}
      multiple
      onchange={handleFolderSelect}
    />
    <input
      bind:this={fileInputRef}
      type="file"
      accept=".md"
      class="hidden"
      onchange={handleFileSelect}
    />
  </div>

  <div class="border-t border-(--chat-border) pt-4">
    <div class="text-[10px] uppercase tracking-widest text-(--chat-text-muted) mb-2">
      {t($locale, "settings.about")}
    </div>
    <p class="text-xs text-(--chat-text-secondary) leading-relaxed">
      {adapter.appName || "This app"} {t($locale, "settings.aboutText")}
    </p>
    {#if isCustom}
      <p class="text-xs text-(--chat-text-muted) leading-relaxed mt-2">
        {t($locale, "settings.customEndpointHint")}
      </p>
    {/if}
    {#if useProxy}
      <p class="text-xs text-(--chat-text-muted) leading-relaxed mt-2">
        {t($locale, "settings.corsAbout")}
      </p>
    {/if}
    <p class="text-[10px] text-(--chat-text-muted) mt-3">
      {adapter.appVersion ? `v${adapter.appVersion}` : ""}
    </p>
  </div>
</div>
