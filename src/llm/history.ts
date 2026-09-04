import type { LlmMessage } from "./provider";

/**
 * Conversation history, keyed by Teams conversation id.
 *
 * IN-MEMORY ONLY. This does not survive a process restart, which on Azure App
 * Service includes every deploy and every idle recycle -- users silently lose
 * context. It is also per-instance, so scaling past a single worker gives the
 * same user a different history depending on which instance takes the request.
 *
 * Before this is real, it needs durable shared storage (Cosmos DB, Azure Table
 * Storage, or Redis) behind these same three functions. Nothing outside this
 * file needs to change when that happens.
 */

// Messages, not exchanges: one question plus its answer is two entries.
const MAX_MESSAGES = 20;

const histories = new Map<string, LlmMessage[]>();

export function getHistory(conversationId: string): LlmMessage[] {
  return histories.get(conversationId) ?? [];
}

export function appendToHistory(conversationId: string, ...messages: LlmMessage[]): void {
  const history = [...getHistory(conversationId), ...messages];

  while (history.length > MAX_MESSAGES) {
    history.shift();
  }
  // The API requires history to open on a user message, so after trimming keep
  // dropping until it does.
  while (history.length > 0 && history[0].role !== "user") {
    history.shift();
  }

  histories.set(conversationId, history);
}

export function clearHistory(conversationId: string): void {
  histories.delete(conversationId);
}
