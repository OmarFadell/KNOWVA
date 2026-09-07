/**
 * Provider-neutral LLM types.
 *
 * Nothing here mirrors a vendor SDK. Each adapter translates at its own
 * boundary, so moving to Azure AI Foundry or Copilot means adding an adapter
 * file -- callers never see a vendor type.
 *
 * Tool calling is now wired through these shapes. It arrived as additions
 * rather than edits, exactly as the earlier version of this file predicted:
 *   - `LlmRequest.tools?: LlmToolDefinition[]`
 *   - `LlmResponse.toolCalls?: LlmToolCall[]`, with `"tool_use"` in LlmStopReason
 *   - a `"tool"` role on LlmMessage carrying one tool result back
 * The agent loop (src/agent) is the only caller that populates any of it;
 * the plain chat path in app.ts still sends `{ system, messages }` and works
 * unchanged.
 */

export type LlmRole = "user" | "assistant" | "tool";

/** A single tool invocation the model asked for. */
export interface LlmToolCall {
  /** Provider-issued id. Must be echoed back on the matching tool result. */
  id: string;
  /** Tool name, matching an LlmToolDefinition.name that was offered. */
  name: string;
  /** Parsed arguments. Shape is the tool's own input schema; validate before use. */
  input: Record<string, unknown>;
}

export interface LlmMessage {
  role: LlmRole;
  /**
   * Text content. For an assistant message that only calls tools this may be
   * empty. For a `"tool"` message it is the tool's result rendered as text.
   */
  content: string;
  /** Present on assistant messages that called tools. */
  toolCalls?: LlmToolCall[];
  /** Required on `"tool"` messages: which LlmToolCall.id this answers. */
  toolCallId?: string;
  /** On `"tool"` messages: the tool failed and `content` explains why. */
  isError?: boolean;
}

/**
 * A tool offered to the model. Uses JSON Schema for the argument shape so it
 * stays vendor-neutral -- the adapter maps it to whatever the provider expects.
 */
export interface LlmToolDefinition {
  name: string;
  description: string;
  /** JSON Schema object (`{ type: "object", properties: {...}, required: [...] }`). */
  inputSchema: Record<string, unknown>;
}

/** Why generation stopped. `"other"` absorbs provider-specific reasons we don't act on. */
export type LlmStopReason = "end_turn" | "tool_use" | "max_tokens" | "refusal" | "other";

export interface LlmUsage {
  inputTokens: number;
  outputTokens: number;
}

export interface LlmRequest {
  system: string;
  /** Full conversation so far. Must start with a user message. */
  messages: LlmMessage[];
  maxTokens?: number;
  /** Tools the model may call this turn. Omit or leave empty for a plain completion. */
  tools?: LlmToolDefinition[];
}

export interface LlmResponse {
  text: string;
  stopReason: LlmStopReason;
  /** Populated when stopReason is `"tool_use"`. */
  toolCalls?: LlmToolCall[];
  usage?: LlmUsage;
}

export interface LlmProvider {
  /** Provider id, matching the LLM_PROVIDER value that selects it. */
  readonly name: string;
  /** Concrete model in use, for logging. */
  readonly model: string;
  complete(request: LlmRequest): Promise<LlmResponse>;
}
