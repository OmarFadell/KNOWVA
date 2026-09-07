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

export interface AgentToolResult {
  /** Rendered as text back to the model. */
  content: string;
  /** The call failed; `content` is a message explaining why, for the model to relay. */
  isError?: boolean;
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
      return { text: response.text, stopReason: response.stopReason, iterations: round };
    }

    messages.push({
      role: "assistant",
      content: response.text,
      toolCalls,
    });

    for (const call of toolCalls) {
      const result = await executeTool(byName, call.name, call.input, logger);
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
