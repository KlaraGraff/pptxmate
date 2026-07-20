import "fake-indexeddb/auto";
import { describe, expect, it, vi } from "vitest";
import { AgentContext } from "../src/context";
import {
  AgentRuntime,
  type RuntimeAdapter,
  type RuntimeState,
} from "../src/runtime";

// Stub localStorage for Node
if (typeof globalThis.localStorage === "undefined") {
  const store: Record<string, string> = {};
  (globalThis as any).localStorage = {
    getItem: (key: string) => store[key] ?? null,
    setItem: (key: string, value: string) => {
      store[key] = value;
    },
    removeItem: (key: string) => {
      delete store[key];
    },
    clear: () => {
      for (const key of Object.keys(store)) delete store[key];
    },
  };
}

let nsCounter = 0;

function freshNamespace() {
  nsCounter++;
  return {
    dbName: `RuntimeTestDB_${nsCounter}`,
    dbVersion: 1,
    localStoragePrefix: `runtime-test-${nsCounter}`,
    documentSettingsPrefix: `runtime-test-${nsCounter}`,
    documentIdSettingsKey: `runtime-test-${nsCounter}-document-id`,
  };
}

function createAdapter(
  overrides: Partial<RuntimeAdapter> = {},
): RuntimeAdapter {
  return {
    tools: [],
    buildSystemPrompt: () => "You are a test assistant.",
    getDocumentId: async () => "test-doc-1",
    storageNamespace: freshNamespace(),
    ...overrides,
  };
}

function createRuntime(overrides: Partial<RuntimeAdapter> = {}): AgentRuntime {
  const adapter = createAdapter(overrides);
  const ctx = new AgentContext({
    namespace: adapter.storageNamespace,
    staticFiles: adapter.staticFiles,
  });
  return new AgentRuntime(adapter, ctx);
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

async function captureCompletedWrite(
  runtime: AgentRuntime,
  toolName: string,
  args: unknown,
) {
  const agent = (runtime as any).agent;
  await (runtime as any).handleAgentEvent({
    type: "tool_execution_start",
    toolCallId: "completed-write",
    toolName,
    args,
  });
  const preflight = await agent.beforeToolCall({
    toolCall: { id: "completed-write", name: toolName },
    args,
  });
  expect(preflight).toBeUndefined();
  await (runtime as any).handleAgentEvent({
    type: "tool_execution_end",
    toolCallId: "completed-write",
    toolName,
    result: { content: [{ type: "text", text: '{"success":true}' }] },
    isError: false,
  });
  (runtime as any).captureRecoveryMutations();
}

describe("AgentRuntime", () => {
  it("getModelsForProvider returns empty array for unknown provider", () => {
    const runtime = createRuntime();
    const models = runtime.getModelsForProvider("nonexistent-provider");
    expect(models).toEqual([]);
    runtime.dispose();
  });

  it("applyConfig sets up agent and updates state", () => {
    const runtime = createRuntime();

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    const state = runtime.getState();
    expect(state.providerConfig).not.toBeNull();
    expect(state.providerConfig!.provider).toBe("openai");
    expect(state.providerConfig!.model).toBe("gpt-4o-mini");
    expect(state.sessionStats.contextWindow).toBeGreaterThan(0);
    expect(state.error).toBeNull();
    runtime.dispose();
  });

  it("applyConfig with custom provider builds custom model", () => {
    const runtime = createRuntime();

    runtime.applyConfig({
      provider: "custom",
      apiKey: "test-key",
      model: "llama3",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
      apiType: "openai-completions",
      customBaseUrl: "http://localhost:11434",
    });

    const state = runtime.getState();
    expect(state.providerConfig).not.toBeNull();
    expect(state.providerConfig!.provider).toBe("custom");
    expect(state.sessionStats.contextWindow).toBe(128000);
    runtime.dispose();
  });

  it("strips runtime-only usage carry before provider calls", async () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;

    const converted = await agent.convertToLlm([
      {
        role: "user",
        content: "hello",
        timestamp: 1,
        _runtimeUsage: {
          inputTokens: 10,
          outputTokens: 0,
          cacheRead: 0,
          cacheWrite: 0,
          totalCost: 0.001,
          lastInputTokens: 10,
        },
      },
    ]);

    expect(converted).toHaveLength(1);
    expect(converted[0]).not.toHaveProperty("_runtimeUsage");
    runtime.dispose();
  });

  it("sendMessage errors when no config", async () => {
    const runtime = createRuntime();
    await runtime.sendMessage("hello");
    const state = runtime.getState();
    expect(state.error).toContain("API key");
    runtime.dispose();
  });

  it("passes the user request to metadata routing and keeps only latest metadata", async () => {
    let metadataRequest:
      | {
          userMessage: string;
          messageCount: number;
          signal?: AbortSignal;
        }
      | undefined;
    const runtime = createRuntime({
      metadataTag: "ppt_context",
      metadataHistory: "latest",
      buildSystemPromptForMessage: (message) => `route:${message}`,
      getDocumentMetadata: async (request) => {
        metadataRequest = request;
        return { metadata: { route: "text" } };
      },
    });

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    const agent = (runtime as any).agent;
    agent.state.messages = [
      {
        role: "user",
        content: [
          {
            type: "text",
            text: '<ppt_context>\n{"route":"layout"}\n</ppt_context>\n\nold request',
          },
        ],
        timestamp: 1,
      },
    ];
    let sentPrompt = "";
    agent.prompt = async (prompt: string) => {
      sentPrompt = prompt;
    };

    await runtime.sendMessage("translate the title");

    expect(metadataRequest).toMatchObject({
      userMessage: "translate the title",
      messageCount: 1,
    });
    expect(metadataRequest?.signal).toBeInstanceOf(AbortSignal);
    expect(metadataRequest?.signal.aborted).toBe(false);
    expect(agent.state.messages[0].content[0].text).toBe("old request");
    expect(agent.state.systemPrompt).toBe("route:translate the title");
    expect(sentPrompt).toContain('<ppt_context>\n{\n  "route": "text"\n}');
    runtime.dispose();
  });

  it("routes tools per message using prior requests from the active session", async () => {
    const fullTool = { name: "full_tool" } as any;
    const textTool = { name: "text_tool" } as any;
    let routedInfo: { priorUserMessages: readonly string[] } | undefined;
    const runtime = createRuntime({
      tools: [fullTool],
      metadataTag: "ppt_context",
      toolsForMessage: (_message, _context, info) => {
        routedInfo = info;
        return [textTool];
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    agent.state.messages = [
      {
        role: "user",
        content:
          '<ppt_context>\n{"route":"layout"}\n</ppt_context>\n\ntranslate slide 2',
        timestamp: 1,
      },
    ];
    let promptToolNames: string[] = [];
    agent.prompt = async () => {
      promptToolNames = agent.state.tools.map((tool: { name: string }) =>
        String(tool.name),
      );
    };

    await runtime.sendMessage("continue with the next slide");

    expect(routedInfo?.priorUserMessages).toEqual(["translate slide 2"]);
    expect(promptToolNames).toEqual(["text_tool"]);
    runtime.dispose();
  });

  it("does not prompt after a metadata read is aborted", async () => {
    const metadata = deferred<{ metadata: object } | null>();
    let metadataSignal: AbortSignal | undefined;
    const runtime = createRuntime({
      getDocumentMetadata: async (request) => {
        metadataSignal = request?.signal;
        return metadata.promise;
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    agent.prompt = vi.fn();

    const send = runtime.sendMessage("read the title");
    await vi.waitFor(() => expect(metadataSignal).toBeDefined());
    runtime.abort();
    expect(metadataSignal?.aborted).toBe(true);
    metadata.resolve({ metadata: { route: "text" } });
    await send;

    expect(agent.prompt).not.toHaveBeenCalled();
    expect(runtime.getState().isStreaming).toBe(false);
    expect(runtime.getState().messages).toHaveLength(0);
    runtime.dispose();
  });

  it("uses a provider config selected while metadata is loading", async () => {
    const metadataStarted = deferred<void>();
    const metadata = deferred<{ metadata: object } | null>();
    const routedTool = { name: "text_tool" } as any;
    const runtime = createRuntime({
      toolsForMessage: () => [routedTool],
      buildSystemPromptForMessage: (message) => `route:${message}`,
      getDocumentMetadata: async () => {
        metadataStarted.resolve(undefined);
        return metadata.promise;
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-old",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const oldAgent = (runtime as any).agent;
    oldAgent.prompt = vi.fn();
    let newAgent: any;
    const newPrompt = vi.fn();
    const originalApplyConfig = runtime.applyConfig.bind(runtime);
    vi.spyOn(runtime, "applyConfig").mockImplementation((config) => {
      originalApplyConfig(config);
      newAgent = (runtime as any).agent;
      newAgent.prompt = newPrompt;
    });

    const send = runtime.sendMessage("translate the title");
    await metadataStarted.promise;
    runtime.setProviderConfig({
      provider: "openai",
      apiKey: "sk-new",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    expect((runtime as any).agent).toBe(oldAgent);
    metadata.resolve({ metadata: { route: "text" } });
    await send;

    expect(oldAgent.prompt).not.toHaveBeenCalled();
    expect(newAgent).not.toBe(oldAgent);
    expect(newPrompt).toHaveBeenCalledWith(
      expect.stringContaining('"route": "text"'),
    );
    expect(newAgent.state.tools).toEqual([routedTool]);
    expect(newAgent.state.systemPrompt).toBe("route:translate the title");
    expect(runtime.getState().providerConfig?.apiKey).toBe("sk-new");
    runtime.dispose();
  });

  it("hides a silent-overflow assistant message before retrying", () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    const contextWindow = runtime.getState().sessionStats.contextWindow;
    const assistantMessage = {
      role: "assistant" as const,
      content: [{ type: "text" as const, text: "partial response" }],
      timestamp: 1,
      stopReason: "stop" as const,
      api: "openai-completions" as const,
      provider: "openai",
      model: "gpt-4o-mini",
      usage: {
        input: contextWindow + 1,
        output: 10,
        cacheRead: 0,
        cacheWrite: 0,
        totalTokens: contextWindow + 11,
        cost: {
          input: 0,
          output: 0,
          cacheRead: 0,
          cacheWrite: 0,
          total: 0,
        },
      },
    };

    (runtime as any).handleAgentEvent({
      type: "message_start",
      message: assistantMessage,
    });
    expect(runtime.getState().messages).toHaveLength(1);

    (runtime as any).handleAgentEvent({
      type: "message_end",
      message: assistantMessage,
    });

    expect(runtime.getState().messages).toHaveLength(0);
    expect((runtime as any).continuationReason).toBe("overflow");
    runtime.dispose();
  });

  it("does not queue an automatic continuation for a length-stopped tool call", async () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const contextWindow = runtime.getState().sessionStats.contextWindow;

    await (runtime as any).handleAgentEvent({
      type: "message_end",
      message: {
        role: "assistant",
        content: [
          {
            type: "toolCall",
            id: "tool-1",
            name: "read_slide_texts",
            arguments: { slide_index: 0 },
          },
        ],
        timestamp: 1,
        stopReason: "length",
        api: "openai-completions",
        provider: "openai",
        model: "gpt-4o-mini",
        usage: {
          input: 100,
          output: 10,
          cacheRead: 0,
          cacheWrite: 0,
          totalTokens: 110,
          cost: {
            input: 0,
            output: 0,
            cacheRead: 0,
            cacheWrite: 0,
            total: 0,
          },
        },
      },
    });

    expect(agent.hasQueuedMessages()).toBe(false);
    expect((runtime as any).continuationNeeded).toBe(false);
    expect(contextWindow).toBeGreaterThan(100);
    runtime.dispose();
  });

  it("clears queued continuations when aborted", () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    agent.followUp({ role: "user", content: "continue", timestamp: 1 });
    expect(agent.hasQueuedMessages()).toBe(true);

    runtime.abort();

    expect(agent.hasQueuedMessages()).toBe(false);
    runtime.dispose();
  });

  it("waits for an aborted agent run to settle before sending again", async () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    agent.state.isStreaming = true;
    const order: string[] = [];
    agent.waitForIdle = vi.fn(async () => {
      order.push("idle");
      agent.state.isStreaming = false;
    });
    agent.prompt = vi.fn(async () => {
      order.push("prompt");
    });

    await runtime.sendMessage("next request");

    expect(order).toEqual(["idle", "prompt"]);
    runtime.dispose();
  });

  it("waits for the old agent to settle before applying a pending provider config", async () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const oldAgent = (runtime as any).agent;
    oldAgent._state.isStreaming = true;
    const order: string[] = [];
    oldAgent.waitForIdle = vi.fn(async () => {
      order.push("idle");
      oldAgent._state.isStreaming = false;
    });

    const originalApplyConfig = runtime.applyConfig.bind(runtime);
    vi.spyOn(runtime, "applyConfig").mockImplementation((config) => {
      order.push("config");
      originalApplyConfig(config);
      const newAgent = (runtime as any).agent;
      newAgent.prompt = vi.fn(async () => {
        order.push("prompt");
      });
    });
    runtime.setProviderConfig({
      provider: "openai",
      apiKey: "sk-next",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    await runtime.sendMessage("continue with the new provider config");

    expect(order).toEqual(["idle", "config", "prompt"]);
    expect((runtime as any).agent).not.toBe(oldAgent);
    runtime.dispose();
  });

  it("does not let a stale clear reset a newer send while waiting for idle", async () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const idle = deferred<void>();
    agent._state.isStreaming = true;
    agent.abort = vi.fn(() => {
      agent._state.isStreaming = false;
    });
    agent.waitForIdle = vi.fn(() => idle.promise);
    const reset = vi.spyOn(agent, "reset");
    agent.prompt = vi.fn(async () => undefined);

    const clear = runtime.clearMessages();
    expect(agent.waitForIdle).toHaveBeenCalledOnce();
    expect(reset).not.toHaveBeenCalled();

    await runtime.sendMessage("new request wins");
    idle.resolve(undefined);
    await clear;

    expect(agent.prompt).toHaveBeenCalledWith("new request wins");
    expect(reset).not.toHaveBeenCalled();
    expect(runtime.getState().messages).toHaveLength(1);
    expect(runtime.getState().messages[0].parts).toEqual([
      { type: "text", text: "new request wins" },
    ]);
    runtime.dispose();
  });

  it("does not retry an overflow when the adapter cannot compact context", async () => {
    const runtime = createRuntime();
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    agent.prompt = async (prompt: string) => {
      agent.state.messages.push(
        { role: "user", content: prompt, timestamp: 1 },
        {
          role: "assistant",
          content: [
            {
              type: "toolCall",
              id: "write-1",
              name: "update_slide_text",
              arguments: {},
            },
          ],
          timestamp: 2,
          usage: {
            input: 50,
            output: 2,
            cacheRead: 5,
            cacheWrite: 0,
            totalTokens: 57,
            cost: {
              input: 0.005,
              output: 0.001,
              cacheRead: 0.0005,
              cacheWrite: 0,
              total: 0.0065,
            },
          },
        },
        {
          role: "toolResult",
          toolCallId: "write-1",
          content: [{ type: "text", text: "unexpected write" }],
          timestamp: 3,
        },
        { role: "assistant", content: [], timestamp: 4 },
      );
      (runtime as any).continuationNeeded = true;
      (runtime as any).continuationReason = "overflow";
    };
    agent.continue = vi.fn();

    await runtime.sendMessage("hello");

    expect(agent.continue).not.toHaveBeenCalled();
    expect(agent.state.messages.at(-1)?.role).toBe("user");
    expect(agent.state.messages).toHaveLength(1);
    expect(agent.state.messages[0]._runtimeUsage).toMatchObject({
      inputTokens: 50,
      outputTokens: 2,
      cacheRead: 5,
      totalCost: 0.0065,
    });
    expect(runtime.getState().sessionStats).toMatchObject({
      inputTokens: 50,
      outputTokens: 2,
      cacheRead: 5,
      totalCost: 0.0065,
    });
    expect(runtime.getState().isStreaming).toBe(false);
    runtime.dispose();
  });

  it("preserves completed writes on overflow and blocks an exact replay", async () => {
    const runtime = createRuntime({
      transformContext: async (messages) => messages,
      toolExecution: "sequential",
      getToolRecoveryInfo: (toolName, args) => ({
        effect: toolName === "update_slide_text" ? "write" : "read",
        mutationKind: toolName === "update_slide_text" ? "text" : undefined,
        verificationKinds: toolName === "read_slide_texts" ? ["text"] : [],
        scope:
          args &&
          typeof args === "object" &&
          typeof (args as { slide_index?: unknown }).slide_index === "number"
            ? {
                slide_index: (args as { slide_index: number }).slide_index,
              }
            : undefined,
      }),
      normalizeToolArgsForReplay: (_toolName, args) => {
        if (!args || typeof args !== "object") return args;
        const normalized = { ...(args as Record<string, unknown>) };
        delete normalized.explanation;
        if (Array.isArray(normalized.updates)) {
          normalized.updates = normalized.updates.map((update) => ({
            ...(update as Record<string, unknown>),
            mode: (update as { mode?: string }).mode ?? "replace",
          }));
        }
        return normalized;
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const contextWindow = runtime.getState().sessionStats.contextWindow;
    const writeArgs = {
      slide_index: 0,
      updates: [{ shape_id: "2", text: "SECRET_TRANSLATION" }],
      explanation: "first explanation",
    };
    const equivalentReplayArgs = {
      slide_index: 0,
      updates: [
        {
          shape_id: "2",
          text: "SECRET_TRANSLATION",
          mode: "replace",
        },
      ],
      explanation: "different explanation",
    };
    const prompts: string[] = [];
    let exactReplayDuringRecovery:
      | { block?: boolean; reason?: string }
      | undefined;
    let differentWriteDuringRecovery:
      | { block?: boolean; reason?: string }
      | undefined;
    agent.prompt = vi.fn(async (prompt: string) => {
      prompts.push(prompt);
      if (prompts.length > 1) {
        exactReplayDuringRecovery = await agent.beforeToolCall({
          toolCall: { id: "replay", name: "update_slide_text" },
          args: equivalentReplayArgs,
        });
        differentWriteDuringRecovery = await agent.beforeToolCall({
          toolCall: { id: "different", name: "update_slide_text" },
          args: { ...writeArgs, slide_index: 1 },
        });
        await (runtime as any).handleAgentEvent({
          type: "message_end",
          message: {
            role: "assistant",
            content: [{ type: "text", text: "done" }],
            timestamp: 5,
            stopReason: "stop",
            api: "openai-completions",
            provider: "openai",
            model: "gpt-4o-mini",
            usage: {
              input: 10,
              output: 2,
              cacheRead: 0,
              cacheWrite: 0,
              totalTokens: 12,
              cost: {
                input: 0,
                output: 0,
                cacheRead: 0,
                cacheWrite: 0,
                total: 0,
              },
            },
          },
        });
        await (runtime as any).handleAgentEvent({
          type: "agent_end",
          messages: [],
        });
        return;
      }
      agent.state.messages.push(
        { role: "user", content: prompt, timestamp: 1 },
        {
          role: "assistant",
          content: [
            {
              type: "toolCall",
              id: "write-complete",
              name: "update_slide_text",
              arguments: writeArgs,
            },
          ],
          timestamp: 2,
        },
        {
          role: "toolResult",
          toolCallId: "write-complete",
          toolName: "update_slide_text",
          content: [{ type: "text", text: '{"success":true}' }],
          isError: false,
          timestamp: 3,
        },
        {
          role: "assistant",
          content: [{ type: "text", text: "" }],
          timestamp: 4,
          stopReason: "stop",
          api: "openai-completions",
          provider: "openai",
          model: "gpt-4o-mini",
          usage: {
            input: contextWindow + 1,
            output: 1,
            cacheRead: 0,
            cacheWrite: 0,
            totalTokens: contextWindow + 2,
            cost: {
              input: 0,
              output: 0,
              cacheRead: 0,
              cacheWrite: 0,
              total: 0,
            },
          },
        },
      );
      await (runtime as any).handleAgentEvent({
        type: "tool_execution_start",
        toolCallId: "write-complete",
        toolName: "update_slide_text",
        args: writeArgs,
      });
      await agent.beforeToolCall({
        toolCall: {
          id: "write-complete",
          name: "update_slide_text",
        },
        args: writeArgs,
      });
      await (runtime as any).handleAgentEvent({
        type: "tool_execution_end",
        toolCallId: "write-complete",
        toolName: "update_slide_text",
        result: {
          content: [{ type: "text", text: '{"success":true}' }],
        },
        isError: false,
      });
      (runtime as any).continuationNeeded = true;
      (runtime as any).continuationReason = "overflow";
    });

    await runtime.sendMessage("translate slide 1");

    expect(prompts).toHaveLength(2);
    expect(prompts[1]).toContain("<runtime_recovery>");
    expect(prompts[1]).toContain("completed text: update_slide_text");
    expect(prompts[1]).toContain("slide_index=0");
    expect(prompts[1]).not.toContain("SECRET_TRANSLATION");
    expect(
      agent.state.messages.some(
        (message: { role: string; toolCallId?: string }) =>
          message.role === "toolResult" &&
          message.toolCallId === "write-complete",
      ),
    ).toBe(true);
    expect(agent.toolExecution).toBe("sequential");
    expect(exactReplayDuringRecovery?.block).toBe(true);
    expect(exactReplayDuringRecovery?.reason).toContain("RUNTIME_REPLAY_GUARD");
    expect(differentWriteDuringRecovery).toBeUndefined();

    const replayAfterRecovery = await agent.beforeToolCall({
      toolCall: { id: "later-replay", name: "update_slide_text" },
      args: equivalentReplayArgs,
    });
    expect(replayAfterRecovery).toBeUndefined();
    runtime.dispose();
  });

  it("keeps completed replay guards when a recovery prompt fails", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: () => ({
        effect: "write",
        mutationKind: "text",
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const args = { slide_index: 0, text: "translated" };
    await captureCompletedWrite(runtime, "update_slide_text", args);
    agent.prompt = vi.fn(async () => {
      throw new Error("recovery failed");
    });

    await runtime.sendMessage("continue recovery");

    expect(agent.prompt).toHaveBeenCalledWith(
      expect.stringContaining("completed text: update_slide_text"),
    );
    expect(runtime.getState().error).toContain("recovery failed");
    const replay = await agent.beforeToolCall({
      toolCall: { id: "replay-after-error", name: "update_slide_text" },
      args,
    });
    expect(replay?.reason).toContain("RUNTIME_REPLAY_GUARD");
    runtime.dispose();
  });

  it("keeps completed replay guards when a recovery prompt is aborted", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: () => ({
        effect: "write",
        mutationKind: "text",
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const args = { slide_index: 0, text: "translated" };
    await captureCompletedWrite(runtime, "update_slide_text", args);
    const promptStarted = deferred<void>();
    const releasePrompt = deferred<void>();
    agent.prompt = vi.fn(async () => {
      promptStarted.resolve(undefined);
      await releasePrompt.promise;
    });

    const send = runtime.sendMessage("continue recovery");
    await promptStarted.promise;
    runtime.abort();
    releasePrompt.resolve(undefined);
    await send;

    expect(agent.prompt).toHaveBeenCalledWith(
      expect.stringContaining("completed text: update_slide_text"),
    );
    const replay = await agent.beforeToolCall({
      toolCall: { id: "replay-after-abort", name: "update_slide_text" },
      args,
    });
    expect(replay?.reason).toContain("RUNTIME_REPLAY_GUARD");
    runtime.dispose();
  });

  it("requires verification after an uncertain write, then permits missing work", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: (toolName, args) => {
        const slideIndex =
          args && typeof args === "object"
            ? (args as { slide_index?: number }).slide_index
            : undefined;
        if (toolName === "read_slide_texts" || toolName === "verify_slides") {
          return {
            effect: "read",
            verificationKinds:
              toolName === "read_slide_texts" ? ["arbitrary"] : ["layout"],
            scope:
              slideIndex === undefined
                ? undefined
                : { slide_index: slideIndex },
          };
        }
        return {
          effect: "unknown",
          mutationKind: "arbitrary",
          scope:
            slideIndex === undefined ? undefined : { slide_index: slideIndex },
        };
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const uncertainArgs = { code: "SECRET_CODE", slide_index: 2 };

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "unknown-write",
      toolName: "execute_office_js",
      args: uncertainArgs,
    });
    await agent.beforeToolCall({
      toolCall: { id: "unknown-write", name: "execute_office_js" },
      args: uncertainArgs,
    });
    (runtime as any).captureRecoveryMutations();

    const blocked = await agent.beforeToolCall({
      toolCall: { id: "blocked-before-read", name: "execute_office_js" },
      args: { code: "different code" },
    });
    expect(blocked?.reason).toContain("VERIFY_FIRST");

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "wrong-slide-read",
      toolName: "read_slide_texts",
      args: { slide_index: 1 },
    });
    await agent.beforeToolCall({
      toolCall: { id: "wrong-slide-read", name: "read_slide_texts" },
      args: { slide_index: 1 },
    });
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "wrong-slide-read",
      toolName: "read_slide_texts",
      result: { content: [{ type: "text", text: '{"success":true}' }] },
      isError: false,
    });
    await (runtime as any).handleAgentEvent({
      type: "turn_end",
      message: { role: "assistant", content: [] },
      toolResults: [],
    });

    const blockedAfterWrongSlide = await agent.beforeToolCall({
      toolCall: { id: "still-blocked", name: "execute_office_js" },
      args: uncertainArgs,
    });
    expect(blockedAfterWrongSlide?.reason).toContain("VERIFY_FIRST");

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "wrong-kind-read",
      toolName: "verify_slides",
      args: { slide_index: 2 },
    });
    await agent.beforeToolCall({
      toolCall: { id: "wrong-kind-read", name: "verify_slides" },
      args: { slide_index: 2 },
    });
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "wrong-kind-read",
      toolName: "verify_slides",
      result: { content: [{ type: "text", text: '{"success":true}' }] },
      isError: false,
    });
    await (runtime as any).handleAgentEvent({
      type: "turn_end",
      message: { role: "assistant", content: [] },
      toolResults: [],
    });

    const blockedAfterWrongKind = await agent.beforeToolCall({
      toolCall: { id: "still-blocked-by-kind", name: "execute_office_js" },
      args: uncertainArgs,
    });
    expect(blockedAfterWrongKind?.reason).toContain("VERIFY_FIRST");

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "verify-read",
      toolName: "read_slide_texts",
      args: { slide_index: 2 },
    });
    await agent.beforeToolCall({
      toolCall: { id: "verify-read", name: "read_slide_texts" },
      args: { slide_index: 2 },
    });
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "verify-read",
      toolName: "read_slide_texts",
      result: { content: [{ type: "text", text: '{"success":true}' }] },
      isError: false,
    });

    const blockedInSameBatch = await agent.beforeToolCall({
      toolCall: { id: "same-batch-write", name: "execute_office_js" },
      args: uncertainArgs,
    });
    expect(blockedInSameBatch?.reason).toContain("VERIFY_FIRST");

    await (runtime as any).handleAgentEvent({
      type: "turn_end",
      message: { role: "assistant", content: [] },
      toolResults: [],
    });

    const allowed = await agent.beforeToolCall({
      toolCall: { id: "allowed-after-read", name: "execute_office_js" },
      args: uncertainArgs,
    });
    expect(allowed).toBeUndefined();
    runtime.dispose();
  });

  it("treats an admitted semantic write error as uncertain", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: (_toolName, args) => ({
        effect: "write",
        mutationKind: "text",
        scope: {
          slide_index: Number(
            (args as { slide_index?: number } | undefined)?.slide_index ?? 0,
          ),
        },
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "failed-write",
      toolName: "update_slide_text",
      args: { slide_index: 0 },
    });
    const agent = (runtime as any).agent;
    const preflight = await agent.beforeToolCall({
      toolCall: { id: "failed-write", name: "update_slide_text" },
      args: { slide_index: 0 },
    });
    expect(preflight).toBeUndefined();
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "failed-write",
      toolName: "update_slide_text",
      result: {
        content: [{ type: "text", text: '{"success":false,"error":"bad id"}' }],
      },
      isError: false,
    });

    expect((runtime as any).recoveryMutations()).toEqual([
      {
        toolCallId: "failed-write",
        toolName: "update_slide_text",
        effect: "write",
        kind: "text",
        status: "uncertain",
        scope: { slide_index: "0" },
      },
    ]);
    const blocked = await agent.beforeToolCall({
      toolCall: { id: "next-write", name: "update_slide_text" },
      args: { slide_index: 1 },
    });
    expect(blocked?.reason).toContain("VERIFY_FIRST");
    runtime.dispose();
  });

  it.each([
    {
      label: "the current mutation state",
      contract: { mutationState: "not_started", mutationCompleted: true },
    },
    {
      label: "the legacy completion flag",
      contract: { mutationCompleted: false },
    },
  ])("does not recover an admitted prewrite failure reported by $label", async ({
    contract,
  }) => {
    const runtime = createRuntime({
      getToolRecoveryInfo: (_toolName, args) => {
        const record =
          args && typeof args === "object"
            ? (args as Record<string, unknown>)
            : {};
        return {
          effect: "write",
          mutationKind: "text",
          scope:
            typeof record.slide_id === "string"
              ? { slide_id: record.slide_id }
              : typeof record.slide_index === "number"
                ? { slide_index: record.slide_index }
                : undefined,
        };
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const args = { slide_index: 7, text: "Bonjour" };

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "safe-prewrite-failure",
      toolName: "update_slide_text",
      args,
    });
    expect(
      await agent.beforeToolCall({
        toolCall: {
          id: "safe-prewrite-failure",
          name: "update_slide_text",
        },
        args,
      }),
    ).toBeUndefined();
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "safe-prewrite-failure",
      toolName: "update_slide_text",
      result: {
        content: [
          {
            type: "text",
            text: [
              "command diagnostic",
              JSON.stringify({
                success: false,
                error: "stale slide directory",
                ...contract,
              }),
              "[exit code: 1]",
            ].join("\n"),
          },
        ],
      },
      isError: false,
    });

    expect((runtime as any).recoveryMutations()).toEqual([]);
    expect((runtime as any).recoveryVerificationRequired).toBe(false);
    const nextWrite = await agent.beforeToolCall({
      toolCall: { id: "write-after-safe-failure", name: "update_slide_text" },
      args: { slide_index: 6, text: "Salut" },
    });
    expect(nextWrite).toBeUndefined();
    runtime.dispose();
  });

  it("uses a successful replacement ID for recovery scope and replay guards", async () => {
    const stableSlideId = (record: Record<string, unknown>) => {
      for (const key of [
        "_modifiedSlideId",
        "replacementSlideId",
        "newSlideId",
        "slide_id",
        "slideId",
      ]) {
        if (typeof record[key] === "string") return record[key] as string;
      }
      return undefined;
    };
    const runtime = createRuntime({
      getToolRecoveryInfo: (_toolName, args) => {
        const record =
          args && typeof args === "object"
            ? (args as Record<string, unknown>)
            : {};
        const slideId = stableSlideId(record);
        const scope: Record<string, string | number | boolean> = {};
        if (slideId) {
          scope.slide_id = slideId;
        } else if (typeof record.slide_index === "number") {
          scope.slide_index = record.slide_index;
        }
        if (typeof record.directoryVersion === "string") {
          scope.directory_version = record.directoryVersion;
        }
        if (typeof record.originalSlideId === "string") {
          scope.original_slide_id = record.originalSlideId;
        }
        if (typeof record.replacementSlideId === "string") {
          scope.replacement_slide_id = record.replacementSlideId;
        }
        return {
          effect: "write",
          mutationKind: "arbitrary",
          scope: Object.keys(scope).length > 0 ? scope : undefined,
        };
      },
      normalizeToolArgsForReplay: (_toolName, args) => {
        if (!args || typeof args !== "object") return args;
        const normalized = { ...(args as Record<string, unknown>) };
        const slideId = stableSlideId(normalized);
        if (slideId) {
          normalized.slide_id = slideId;
          delete normalized._modifiedSlideId;
          delete normalized.replacementSlideId;
          delete normalized.newSlideId;
          delete normalized.slideId;
          delete normalized.slide_index;
          delete normalized.directoryVersion;
        }
        delete normalized.explanation;
        return normalized;
      },
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const originalArgs = {
      slide_id: "slide-H",
      slide_index: 7,
      directoryVersion: "directory-v1",
      code: "return markDirty();",
      explanation: "update original slide 8",
    };

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "replacement-write",
      toolName: "edit_slide_xml",
      args: originalArgs,
    });
    expect(
      await agent.beforeToolCall({
        toolCall: { id: "replacement-write", name: "edit_slide_xml" },
        args: originalArgs,
      }),
    ).toBeUndefined();
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "replacement-write",
      toolName: "edit_slide_xml",
      result: {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              success: true,
              _modifiedSlideId: "slide-H-replacement",
              originalSlideId: "slide-H",
              replacementSlideId: "slide-H-replacement",
              slideIndex: 6,
              directoryVersion: "directory-v2",
            }),
          },
        ],
      },
      isError: false,
    });
    (runtime as any).captureRecoveryMutations();

    expect((runtime as any).recoveryMutations()).toEqual([
      {
        toolCallId: "replacement-write",
        toolName: "edit_slide_xml",
        effect: "write",
        kind: "arbitrary",
        status: "completed",
        scope: {
          slide_id: "slide-H-replacement",
          directory_version: "directory-v2",
          original_slide_id: "slide-H",
          replacement_slide_id: "slide-H-replacement",
        },
      },
    ]);
    const originalReplay = await agent.beforeToolCall({
      toolCall: { id: "replay-old-id", name: "edit_slide_xml" },
      args: originalArgs,
    });
    const replacementReplay = await agent.beforeToolCall({
      toolCall: { id: "replay-new-id", name: "edit_slide_xml" },
      args: {
        ...originalArgs,
        slide_id: "slide-H-replacement",
        slide_index: 6,
        directoryVersion: "directory-v2",
        explanation: "retry against the replacement",
      },
    });
    expect(originalReplay?.reason).toContain("RUNTIME_REPLAY_GUARD");
    expect(replacementReplay?.reason).toContain("RUNTIME_REPLAY_GUARD");
    runtime.dispose();
  });

  it("does not let a broad slide read verify a shape-scoped mutation", () => {
    const runtime = createRuntime();
    const mutation = {
      toolCallId: "shape-write",
      toolName: "edit_slide_text",
      effect: "write",
      kind: "text",
      status: "uncertain",
      scope: { slide_index: "0", shape_id: "7" },
    };

    expect(
      (runtime as any).verificationCoversMutation(
        {
          effect: "read",
          verificationKinds: ["text"],
          scope: { slide_index: "0" },
        },
        mutation,
      ),
    ).toBe(false);
    expect(
      (runtime as any).verificationCoversMutation(
        {
          effect: "read",
          verificationKinds: ["text"],
          scope: { slide_index: "0", shape_id: "7" },
        },
        mutation,
      ),
    ).toBe(true);
    runtime.dispose();
  });

  it("does not journal a write rejected by recovery preflight", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: () => ({
        effect: "write",
        mutationKind: "text",
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "uncertain-write",
      toolName: "update_slide_text",
      args: { slide_index: 0 },
    });
    await agent.beforeToolCall({
      toolCall: { id: "uncertain-write", name: "update_slide_text" },
      args: { slide_index: 0 },
    });
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "uncertain-write",
      toolName: "update_slide_text",
      result: { content: [{ type: "text", text: "write failed" }] },
      isError: true,
    });

    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "preflight-blocked",
      toolName: "update_slide_text",
      args: { slide_index: 1 },
    });
    const blocked = await agent.beforeToolCall({
      toolCall: { id: "preflight-blocked", name: "update_slide_text" },
      args: { slide_index: 1 },
    });
    expect(blocked?.reason).toContain("VERIFY_FIRST");
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "preflight-blocked",
      toolName: "update_slide_text",
      result: { content: [{ type: "text", text: "blocked" }] },
      isError: true,
    });

    expect(
      (runtime as any)
        .recoveryMutations()
        .map((mutation: { toolCallId: string }) => mutation.toolCallId),
    ).toEqual(["uncertain-write"]);
    runtime.dispose();
  });

  it("does not admit a write when tool preflight starts with an aborted signal", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: () => ({
        effect: "write",
        mutationKind: "text",
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const args = { slide_index: 0 };
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "aborted-before-entry",
      toolName: "update_slide_text",
      args,
    });
    const controller = new AbortController();
    controller.abort();
    await agent.beforeToolCall(
      {
        toolCall: { id: "aborted-before-entry", name: "update_slide_text" },
        args,
      },
      controller.signal,
    );
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_end",
      toolCallId: "aborted-before-entry",
      toolName: "update_slide_text",
      result: { content: [{ type: "text", text: "aborted" }] },
      isError: true,
    });

    expect((runtime as any).recoveryMutations()).toEqual([]);
    runtime.dispose();
  });

  it("injects compact recovery context on the next message after manual abort", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: (_toolName, args) => ({
        effect: "unknown",
        mutationKind: "arbitrary",
        scope:
          args && typeof args === "object"
            ? {
                slide_index: Number(
                  (args as { slide_index?: number }).slide_index ?? 0,
                ),
              }
            : undefined,
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const codeArgs = { slide_index: 4, code: "SECRET_ABORTED_CODE" };
    await (runtime as any).handleAgentEvent({
      type: "tool_execution_start",
      toolCallId: "active-write",
      toolName: "execute_office_js",
      args: codeArgs,
    });
    const preflight = await agent.beforeToolCall({
      toolCall: { id: "active-write", name: "execute_office_js" },
      args: codeArgs,
    });
    expect(preflight).toBeUndefined();
    agent._state.isStreaming = true;
    agent.abort = vi.fn();

    runtime.abort();
    agent._state.isStreaming = false;
    let sentPrompt = "";
    agent.prompt = vi.fn(async (prompt: string) => {
      sentPrompt = prompt;
      await (runtime as any).handleAgentEvent({
        type: "agent_end",
        messages: [],
      });
    });

    await runtime.sendMessage("continue carefully");

    expect(sentPrompt).toContain("<runtime_recovery>");
    expect(sentPrompt).toContain("uncertain arbitrary: execute_office_js");
    expect(sentPrompt).toContain("slide_index=4");
    expect(sentPrompt).toContain("continue carefully");
    expect(sentPrompt).not.toContain("SECRET_ABORTED_CODE");
    runtime.dispose();
  });

  it("retains a completed write receipt after the journal fills with reads", async () => {
    const runtime = createRuntime({
      getToolRecoveryInfo: (toolName, args) => ({
        effect: toolName === "update_slide_text" ? "write" : "read",
        mutationKind: toolName === "update_slide_text" ? "text" : undefined,
        verificationKinds: [],
        scope:
          args && typeof args === "object"
            ? {
                slide_index: Number(
                  (args as { slide_index?: number }).slide_index ?? 0,
                ),
              }
            : undefined,
      }),
    });
    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });
    const agent = (runtime as any).agent;
    const writeArgs = { slide_index: 3, text: "translated" };
    (runtime as any).recordToolStart(
      "durable-write",
      "update_slide_text",
      writeArgs,
    );
    await agent.beforeToolCall({
      toolCall: { id: "durable-write", name: "update_slide_text" },
      args: writeArgs,
    });
    (runtime as any).recordToolEnd(
      "durable-write",
      "update_slide_text",
      writeArgs,
      { content: [{ type: "text", text: '{"success":true}' }] },
      false,
    );

    for (let index = 0; index < 130; index++) {
      const toolCallId = `read-${index}`;
      const readArgs = { slide_index: index };
      (runtime as any).recordToolStart(
        toolCallId,
        "read_slide_texts",
        readArgs,
      );
      (runtime as any).recordToolEnd(
        toolCallId,
        "read_slide_texts",
        readArgs,
        { content: [{ type: "text", text: '{"success":true}' }] },
        false,
      );
    }

    expect((runtime as any).toolJournal.size).toBe(128);
    expect((runtime as any).toolJournal.has("durable-write")).toBe(true);
    agent._state.isStreaming = true;
    agent.abort = vi.fn();
    runtime.abort();
    agent._state.isStreaming = false;
    let recoveryPrompt = "";
    let guardedReplay: { block?: boolean; reason?: string } | undefined;
    agent.prompt = vi.fn(async (prompt: string) => {
      recoveryPrompt = prompt;
      guardedReplay = await agent.beforeToolCall({
        toolCall: { id: "replay-durable-write", name: "update_slide_text" },
        args: writeArgs,
      });
    });

    await runtime.sendMessage("resume after journal pressure");

    expect(recoveryPrompt).toContain("completed text: update_slide_text");
    expect(recoveryPrompt).toContain("slide_index=3");
    expect(guardedReplay?.reason).toContain("RUNTIME_REPLAY_GUARD");
    runtime.dispose();
  });

  it("clearMessages resets state", async () => {
    const runtime = createRuntime();

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    await runtime.clearMessages();
    const state = runtime.getState();
    expect(state.messages).toEqual([]);
    expect(state.error).toBeNull();
    expect(state.uploads).toEqual([]);
    runtime.dispose();
  });

  it("toggleFollowMode flips followMode", () => {
    const runtime = createRuntime();

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    expect(runtime.getState().providerConfig!.followMode).toBe(true);
    runtime.toggleFollowMode();
    expect(runtime.getState().providerConfig!.followMode).toBe(false);
    runtime.toggleFollowMode();
    expect(runtime.getState().providerConfig!.followMode).toBe(true);
    runtime.dispose();
  });

  it("toggleExpandToolCalls flips expandToolCalls", () => {
    const runtime = createRuntime();

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    expect(runtime.getState().providerConfig!.expandToolCalls).toBe(false);
    runtime.toggleExpandToolCalls();
    expect(runtime.getState().providerConfig!.expandToolCalls).toBe(true);
    runtime.dispose();
  });

  it("uploadFiles adds files and updates state", async () => {
    const runtime = createRuntime();
    await runtime.init();

    await runtime.uploadFiles([
      {
        name: "data.csv",
        size: 100,
        data: new TextEncoder().encode("a,b\n1,2"),
      },
    ]);

    const state = runtime.getState();
    expect(state.uploads).toHaveLength(1);
    expect(state.uploads[0].name).toBe("data.csv");
    expect(state.isUploading).toBe(false);
    runtime.dispose();
  });

  it("keeps upload state locked until queued uploads finish", async () => {
    const runtime = createRuntime();
    await runtime.init();
    const originalWrite = runtime.context.writeFile.bind(runtime.context);
    let releaseFirstWrite: (() => void) | undefined;
    let releaseSecondWrite: (() => void) | undefined;
    const firstWriteBlocked = new Promise<void>((resolve) => {
      releaseFirstWrite = resolve;
    });
    const secondWriteBlocked = new Promise<void>((resolve) => {
      releaseSecondWrite = resolve;
    });
    let writeCount = 0;
    vi.spyOn(runtime.context, "writeFile").mockImplementation(
      async (...args) => {
        writeCount++;
        if (writeCount === 1) await firstWriteBlocked;
        if (writeCount === 2) await secondWriteBlocked;
        return originalWrite(...args);
      },
    );

    const first = runtime.uploadFiles([
      { name: "first.png", size: 1, data: new Uint8Array([1]) },
    ]);
    const second = runtime.uploadFiles([
      { name: "second.png", size: 1, data: new Uint8Array([2]) },
    ]);

    expect(runtime.getState().isUploading).toBe(true);
    releaseFirstWrite?.();
    await first;
    expect(runtime.getState().isUploading).toBe(true);
    releaseSecondWrite?.();
    await second;
    expect(runtime.getState().isUploading).toBe(false);
    expect(runtime.getState().uploads.map((file) => file.name)).toEqual([
      "first.png",
      "second.png",
    ]);
    runtime.dispose();
  });

  it("shows upload failures in runtime state", async () => {
    const runtime = createRuntime();
    await runtime.init();
    vi.spyOn(runtime.context, "writeFile").mockRejectedValueOnce(
      new Error("disk full"),
    );

    await runtime.uploadFiles([
      { name: "failed.png", size: 1, data: new Uint8Array([1]) },
    ]);

    expect(runtime.getState().isUploading).toBe(false);
    expect(runtime.getState().error).toBe("Failed to upload file: disk full");
    expect(runtime.getState().uploads).toEqual([]);
    runtime.dispose();
  });

  it("removeUpload removes a file from state", async () => {
    const runtime = createRuntime();
    await runtime.init();

    await runtime.uploadFiles([
      {
        name: "temp.txt",
        size: 10,
        data: new TextEncoder().encode("temp"),
      },
    ]);
    expect(runtime.getState().uploads).toHaveLength(1);

    await runtime.removeUpload("temp.txt");
    expect(runtime.getState().uploads).toHaveLength(0);
    runtime.dispose();
  });

  it("init loads session and skills", async () => {
    const runtime = createRuntime();
    await runtime.init();

    const state = runtime.getState();
    expect(state.currentSession).not.toBeNull();
    expect(state.currentSession!.workbookId).toBe("test-doc-1");
    expect(Array.isArray(state.skills)).toBe(true);
    runtime.dispose();
  });

  it("init is idempotent", async () => {
    const runtime = createRuntime();
    await runtime.init();
    const session1 = runtime.getState().currentSession;
    await runtime.init();
    const session2 = runtime.getState().currentSession;
    expect(session1!.id).toBe(session2!.id);
    runtime.dispose();
  });

  it("newSession creates a fresh session", async () => {
    const runtime = createRuntime();
    await runtime.init();

    const firstSession = runtime.getState().currentSession!.id;
    await runtime.newSession();
    const secondSession = runtime.getState().currentSession!.id;

    expect(firstSession).not.toBe(secondSession);
    expect(runtime.getState().messages).toEqual([]);
    runtime.dispose();
  });

  it("switchSession restores a previous session", async () => {
    const runtime = createRuntime();
    await runtime.init();

    const firstId = runtime.getState().currentSession!.id;
    await runtime.newSession();
    const secondId = runtime.getState().currentSession!.id;

    await runtime.switchSession(firstId);
    expect(runtime.getState().currentSession!.id).toBe(firstId);

    await runtime.switchSession(secondId);
    expect(runtime.getState().currentSession!.id).toBe(secondId);
    runtime.dispose();
  });

  it("deleteCurrentSession switches to another session", async () => {
    const runtime = createRuntime();
    await runtime.init();

    await runtime.newSession();

    await runtime.deleteCurrentSession();
    expect(runtime.getState().currentSession).not.toBeNull();
    runtime.dispose();
  });

  it("emits state to subscribers on update", async () => {
    const runtime = createRuntime();
    const states: RuntimeState[] = [];
    const unsub = runtime.subscribe((s) => states.push(s));

    runtime.applyConfig({
      provider: "openai",
      apiKey: "sk-test",
      model: "gpt-4o-mini",
      useProxy: false,
      proxyUrl: "",
      thinking: "none",
      followMode: true,
      expandToolCalls: false,
    });

    expect(states.length).toBeGreaterThan(0);
    expect(states[states.length - 1].providerConfig).not.toBeNull();

    unsub();
    runtime.dispose();
  });

  it("init applies adapter.staticFiles to context", async () => {
    const ns = freshNamespace();
    const ctx = new AgentContext({ namespace: ns });

    const adapter = createAdapter({
      storageNamespace: ns,
      staticFiles: {
        "/home/user/docs/word-api.d.ts": "declare const Word: any;",
      },
    });

    expect(await ctx.fileExists("/home/user/docs/word-api.d.ts")).toBe(false);

    const runtime = new AgentRuntime(adapter, ctx);
    await runtime.init();

    expect(await ctx.fileExists("/home/user/docs/word-api.d.ts")).toBe(true);
    expect(await ctx.readFile("/home/user/docs/word-api.d.ts")).toBe(
      "declare const Word: any;",
    );
    runtime.dispose();
  });

  it("init applies adapter.customCommands to context", async () => {
    const ns = freshNamespace();
    const ctx = new AgentContext({ namespace: ns });
    expect(ctx.commandSnippets).toEqual([]);

    const adapter = createAdapter({
      storageNamespace: ns,
      customCommands: () => ({
        commands: [],
        promptSnippets: ["Use `my-cmd` to do stuff"],
      }),
    });

    const runtime = new AgentRuntime(adapter, ctx);
    await runtime.init();

    expect(ctx.commandSnippets).toEqual(["Use `my-cmd` to do stuff"]);
    runtime.dispose();
  });

  it("uploadFiles replaces existing upload with same name", async () => {
    const runtime = createRuntime();
    await runtime.init();

    await runtime.uploadFiles([
      { name: "file.txt", size: 10, data: new TextEncoder().encode("v1") },
    ]);
    await runtime.uploadFiles([
      { name: "file.txt", size: 20, data: new TextEncoder().encode("v2") },
    ]);

    const state = runtime.getState();
    expect(state.uploads).toHaveLength(1);
    expect(state.uploads[0].size).toBe(20);
    runtime.dispose();
  });
});
