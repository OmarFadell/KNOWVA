/**
 * Teams deep links.
 *
 * `chatMessage.webUrl` is null for chat messages -- Graph only populates it for
 * *channel* messages. So a citation that lets someone open the original has to
 * be built by hand from the chat id and message id.
 *
 * Format:
 *   https://teams.microsoft.com/l/message/{chatId}/{messageId}
 *     ?tenantId={tenantId}&context={"contextType":"chat"}
 *
 * Docs:
 *   https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/deep-link-teams
 */

/**
 * Deep link to one message inside a chat.
 *
 * `tenantId` is optional but worth passing: without it, a user signed into more
 * than one tenant can land in the wrong one and see "conversation not found".
 */
export function chatMessageDeepLink(
  chatId: string,
  messageId: string,
  tenantId?: string
): string {
  const context = encodeURIComponent(JSON.stringify({ contextType: "chat" }));
  const tenant = tenantId ? `tenantId=${encodeURIComponent(tenantId)}&` : "";

  // The chat id contains ':' and '@', both of which must survive as a single
  // path segment rather than being read as URL structure.
  return (
    `https://teams.microsoft.com/l/message/${encodeURIComponent(chatId)}/` +
    `${encodeURIComponent(messageId)}?${tenant}context=${context}`
  );
}
