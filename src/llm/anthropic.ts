import Anthropic from "@anthropic-ai/sdk";
import type {
  LlmMessage,
  LlmProvider,
  LlmRequest,
  LlmResponse,
  LlmStopReason,
  LlmToolCall,
  LlmToolDefinition,
} from "./provider";

// This file is the only place in the codebase allowed to import
// @anthropic-ai/sdk. If an Anthropic type escapes past this boundary, the
// provider abstraction is not doing its job.

const DEFAULT_MAX_TOKENS = 8192;

export interface AnthropicProviderOptions {
  apiKey: string;
  model: string;
}

export class AnthropicProvider implements LlmProvider {
  readonly name = "anthropic";
  readonly model: string;
  private readonly client: Anthropic;

  constructor(options: AnthropicProviderOptions) {
    this.model = options.model;
    this.client = new Anthropic({ apiKey: options.apiKey });
  }

  async complete(request: LlmRequest): Promise<LlmResponse> {
    // Plain (non-beta) endpoint deliberately: `model` is configurable via
    // ANTHROPIC_MODEL, and beta request shapes (e.g. the refusal-fallbacks
    // `fallbacks` param) are only supported on specific models -- sending one
    // to a model that doesn't support it is a hard 400, not a graceful no-op.
    // Stick to the stable request surface so any current model works. Tool
    // calling is part of that stable surface -- no beta header needed.
    const response = await this.client.messages.create({
      model: this.model,
      max_tokens: request.maxTokens ?? DEFAULT_MAX_TOKENS,
      system: request.system,
      messages: toAnthropicMessages(request.messages),
      // Empty `tools: []` is a 400, so only send the field when we have some.
      ...(request.tools && request.tools.length > 0
        ? { tools: request.tools.map(toAnthropicTool) }
        : {}),
      // A Teams reply is latency-sensitive and the one tool we orchestrate
      // (SharePoint search) needs at most a couple of rounds, so the cheapest
      // effort level is still the right default. Thinking itself stays on
      // (adaptive is the default on current models) -- explicitly disabling it
      // has known failure modes.
      output_config: { effort: "low" },
    });

    return {
      text: extractText(response.content),
      stopReason: toLlmStopReason(response.stop_reason),
      toolCalls: extractToolCalls(response.content),
      usage: {
        inputTokens: response.usage.input_tokens,
        outputTokens: response.usage.output_tokens,
      },
    };
  }
}

function toAnthropicTool(tool: LlmToolDefinition): Anthropic.Tool {
  return {
    name: tool.name,
    description: tool.description,
    // Our LlmToolDefinition already carries a JSON Schema object; Anthropic's
    // input_schema is exactly that.
    input_schema: tool.inputSchema as Anthropic.Tool.InputSchema,
  };
}

/**
 * Maps the provider-neutral message list to Anthropic's shape.
 *
 * The one non-trivial bit: a `"tool"` message becomes a `user` turn holding a
 * single `tool_result` block. Consecutive tool messages (parallel tool calls in
 * one round) are merged into one user turn -- splitting tool results across
 * multiple user messages trains the model to stop making parallel calls.
 */
function toAnthropicMessages(messages: LlmMessage[]): Anthropic.MessageParam[] {
  const out: Anthropic.MessageParam[] = [];

  for (const message of messages) {
    if (message.role === "tool") {
      const block: Anthropic.ToolResultBlockParam = {
        type: "tool_result",
        tool_use_id: message.toolCallId ?? "",
        content: message.content,
        ...(message.isError ? { is_error: true } : {}),
      };
      const last = out[out.length - 1];
      if (last && last.role === "user" && Array.isArray(last.content)) {
        last.content.push(block);
      } else {
        out.push({ role: "user", content: [block] });
      }
      continue;
    }

    if (message.role === "assistant") {
      const blocks: Anthropic.ContentBlockParam[] = [];
      if (message.content) blocks.push({ type: "text", text: message.content });
      for (const call of message.toolCalls ?? []) {
        blocks.push({
          type: "tool_use",
          id: call.id,
          name: call.name,
          input: call.input,
        });
      }
      out.push({ role: "assistant", content: blocks.length > 0 ? blocks : message.content });
      continue;
    }

    out.push({ role: "user", content: message.content });
  }

  return out;
}

/**
 * Responses carry thinking blocks alongside text blocks. Only text is meant for
 * the user, so anything else is dropped here rather than leaking into a reply.
 */
function extractText(content: Anthropic.ContentBlock[]): string {
  return content
    .filter((block): block is Anthropic.TextBlock => block.type === "text")
    .map((block) => block.text)
    .join("\n\n")
    .trim();
}

function extractToolCalls(content: Anthropic.ContentBlock[]): LlmToolCall[] | undefined {
  const calls = content
    .filter((block): block is Anthropic.ToolUseBlock => block.type === "tool_use")
    .map((block) => ({
      id: block.id,
      name: block.name,
      // block.input is typed `unknown`; every tool we define takes an object.
      input: (block.input ?? {}) as Record<string, unknown>,
    }));
  return calls.length > 0 ? calls : undefined;
}

function toLlmStopReason(stopReason: string | null): LlmStopReason {
  switch (stopReason) {
    case "end_turn":
      return "end_turn";
    case "tool_use":
      return "tool_use";
    case "max_tokens":
      return "max_tokens";
    case "refusal":
      return "refusal";
    default:
      return "other";
  }
}
