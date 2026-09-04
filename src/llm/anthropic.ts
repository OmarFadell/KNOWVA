import Anthropic from "@anthropic-ai/sdk";
import type {
  LlmMessage,
  LlmProvider,
  LlmRequest,
  LlmResponse,
  LlmStopReason,
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
    // Stick to the stable request surface so any current model works.
    const response = await this.client.messages.create({
      model: this.model,
      max_tokens: request.maxTokens ?? DEFAULT_MAX_TOKENS,
      system: request.system,
      messages: request.messages.map(toAnthropicMessage),
      // A Teams reply is latency-sensitive and this milestone has no tools to
      // orchestrate, so the cheapest effort level is the right default. Thinking
      // itself stays on (adaptive is the default on current models) --
      // explicitly disabling it has known failure modes.
      output_config: { effort: "low" },
    });

    return {
      text: extractText(response.content),
      stopReason: toLlmStopReason(response.stop_reason),
      usage: {
        inputTokens: response.usage.input_tokens,
        outputTokens: response.usage.output_tokens,
      },
    };
  }
}

function toAnthropicMessage(message: LlmMessage): Anthropic.MessageParam {
  return { role: message.role, content: message.content };
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

function toLlmStopReason(stopReason: string | null): LlmStopReason {
  switch (stopReason) {
    case "end_turn":
      return "end_turn";
    case "max_tokens":
      return "max_tokens";
    case "refusal":
      return "refusal";
    default:
      return "other";
  }
}
