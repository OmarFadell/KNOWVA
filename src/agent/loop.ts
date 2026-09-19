import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";
import type {
  LlmMessage,
  LlmProvider,
  LlmStopReason,
  LlmToolDefinition,
} from "../llm/provider";

/**
 * The provider-agnostic agent loop: call the model, run any tools it asks for,
 * feed the results back, repeat until it answers with text.
 *
 * Nothing in this file imports a vendor SDK. It talks only to the LlmProvider
 * interface, so it works the same whichever provider LLM_PROVIDER selects.
 */

/**
 * A side-channel from a tool to the *response layer*, for the rare case where
 * text back to the model is not enough on its own.
 *
 * ---------------------------------------------------------------------------
 * WHY THIS EXISTS, SINCE IT IS THE FIRST OF ITS KIND.
 *
 * Every tool so far communicates in exactly one direction: it returns prose,
 * the model reads it, the model writes a reply. "No results found" and "access
 * denied" are handled that way and it works, because in both cases the right
 * outcome IS just words -- the model has everything it needs to say something
 * true.
 *
 * GitHub sign-in breaks that, and it breaks it for a reason worth stating: the
 * right outcome is not words, it is a button. Knowva has to send an Adaptive
 * Card carrying a one-time authorize URL, and a language model cannot be handed
 * a credential-bearing URL and trusted to reproduce it character-perfect in
 * prose -- nor should it be, since that URL is single-use and tied to one user.
 *
 * So the tool returns BOTH: prose telling the model to explain that a
 * connection is needed, and this signal telling the response layer to attach
 * the card. The model never sees the URL.
 *
 * Keep this narrow. It is not a general escape hatch for tools that want to
 * drive the UI; a tool that can express itself in text should.
 * ---------------------------------------------------------------------------
 */
export type AgentToolSignal = {
  /** The user must authorize an external service before this tool can work. */
  kind: "needs-auth";
  /** Which service, so the response layer knows which sign-in to offer. */
  provider: "github";
};

export interface AgentToolResult {
  /** Rendered as text back to the model. */
  content: string;
  /** The call failed; `content` is a message explaining why, for the model to relay. */
  isError?: boolean;
  /**
   * Out-of-band instruction for the response layer. Never shown to the model.
   * See AgentToolSignal.
   */
  signal?: AgentToolSignal;
}

export interface AgentTool {
  readonly definition: LlmToolDefinition;
  run(input: Record<string, unknown>): Promise<AgentToolResult>;
}

export interface RunAgentOptions {
  provider: LlmProvider;
  system: string;
  /** Conversation so far (history + the new user message). Must start with a user message. */
  messages: LlmMessage[];
  tools: AgentTool[];
  logger: ILogger;
  /** Model <-> tool round trips before we force a text answer. 5 is plenty. */
  maxIterations?: number;
  maxTokens?: number;
}

export interface AgentRunResult {
  text: string;
  stopReason: LlmStopReason;
  /** Model calls made, including the final text turn. */
  iterations: number;
  /**
   * Signals raised by tools during the run, in the order they occurred and
   * de-duplicated by kind+provider -- a model that calls search_github twice
   * should not produce two sign-in cards.
   */
  signals: AgentToolSignal[];
}

const DEFAULT_MAX_ITERATIONS = 5;

export async function runAgent(options: RunAgentOptions): Promise<AgentRunResult> {
  const { provider, system, tools, logger } = options;
  const maxIterations = options.maxIterations ?? DEFAULT_MAX_ITERATIONS;

  // Working copy -- the caller's array (and the persisted history) must not gain
  // the intermediate tool turns.
  const messages: LlmMessage[] = [...options.messages];
  const toolDefinitions = tools.map((tool) => tool.definition);
  const byName = new Map(tools.map((tool) => [tool.definition.name, tool] as const));

  // Keyed so a repeated tool call cannot produce a repeated card. Insertion
  // order is preserved, which is what the response layer wants if a turn ever
  // raises more than one kind of signal.
  const signals = new Map<string, AgentToolSignal>();

  for (let round = 1; round <= maxIterations; round++) {
    const response = await provider.complete({
      system,
      messages,
      tools: toolDefinitions,
      maxTokens: options.maxTokens,
    });

    const toolCalls = response.toolCalls ?? [];
    logger.info(
      `agent: round ${round}/${maxIterations} on ${provider.name}/${provider.model} -- ` +
        `stopReason=${response.stopReason} toolCalls=[${toolCalls.map((c) => c.name).join(", ")}] ` +
        `tokens=${response.usage?.inputTokens ?? "?"}in/${response.usage?.outputTokens ?? "?"}out`
    );

    if (response.stopReason !== "tool_use" || toolCalls.length === 0) {
      return {
        text: response.text,
        stopReason: response.stopReason,
        iterations: round,
        signals: [...signals.values()],
      };
    }

    messages.push({
      role: "assistant",
      content: response.text,
      toolCalls,
    });

    for (const call of toolCalls) {
      const result = await executeTool(byName, call.name, call.input, logger);

      if (result.signal) {
        signals.set(`${result.signal.kind}:${result.signal.provider}`, result.signal);
      }

      // Note what is NOT forwarded: `signal`. It is for the response layer
      // only, and the model is given the tool's prose exactly as before.
      messages.push({
        role: "tool",
        toolCallId: call.id,
        content: result.content,
        isError: result.isError,
      });
    }
  }

  // Iteration cap hit and the model still wants tools. Ask once more with no
  // tools offered so it has to commit to a text answer rather than leaving the
  // user with nothing.
  logger.warn(
    `agent: hit the ${maxIterations}-round cap still wanting tools; forcing a final text answer`
  );
  const final = await provider.complete({ system, messages, maxTokens: options.maxTokens });
  return {
    text: final.text,
    stopReason: final.stopReason,
    iterations: maxIterations + 1,
    // Signals raised before the cap still stand: a user who needs to connect
    // GitHub needs to connect it whether or not the model then talked itself
    // out of iterations.
    signals: [...signals.values()],
  };
}

async function executeTool(
  byName: Map<string, AgentTool>,
  name: string,
  input: Record<string, unknown>,
  logger: ILogger
): Promise<AgentToolResult> {
  const tool = byName.get(name);
  if (!tool) {
    // The model invented a tool name. Tell it so, rather than failing the turn.
    logger.warn(`agent: model called unknown tool "${name}"`);
    return { content: `No tool named "${name}" is available.`, isError: true };
  }

  try {
    return await tool.run(input);
  } catch (err) {
    // A tool that throws is a bug in the tool -- it is supposed to turn its own
    // expected failures into an isError result. Log it and hand the model a
    // generic message so the user still gets a reply.
    logger.error(`agent: tool "${name}" threw`, describeError(err), err);
    return {
      content: `The ${name} tool failed unexpectedly. Tell the user it's a problem on our side.`,
      isError: true,
    };
  }
}
