/**
 * Provider-neutral LLM types.
 *
 * Nothing here mirrors a vendor SDK. Each adapter translates at its own
 * boundary, so moving to Azure AI Foundry or Copilot means adding an adapter
 * file -- callers never see a vendor type.
 *
 * Tool calling is deliberately not implemented yet, but the shapes below are
 * chosen so it arrives as additions rather than edits:
 *   - `LlmRequest.tools?: LlmToolDefinition[]`
 *   - `LlmResponse.toolCalls?: LlmToolCall[]`, with `"tool_use"` joining LlmStopReason
 *   - a `"tool"` role on LlmMessage to carry results back
 * Requests and responses are objects rather than positional arguments precisely
 * so each of those is a new optional field and existing call sites keep working.
 */

export type LlmRole = "user" | "assistant";

export interface LlmMessage {
  role: LlmRole;
  content: string;
}

/** Why generation stopped. `"other"` absorbs provider-specific reasons we don't act on. */
export type LlmStopReason = "end_turn" | "max_tokens" | "refusal" | "other";

export interface LlmUsage {
  inputTokens: number;
  outputTokens: number;
}

export interface LlmRequest {
  system: string;
  /** Full conversation so far. Must start with a user message. */
  messages: LlmMessage[];
  maxTokens?: number;
}

export interface LlmResponse {
  text: string;
  stopReason: LlmStopReason;
  usage?: LlmUsage;
}

export interface LlmProvider {
  /** Provider id, matching the LLM_PROVIDER value that selects it. */
  readonly name: string;
  /** Concrete model in use, for logging. */
  readonly model: string;
  complete(request: LlmRequest): Promise<LlmResponse>;
}
