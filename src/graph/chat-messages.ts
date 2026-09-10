import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";

/**
 * Reads messages from one chat, using the app-only token that RSC scopes.
 *
 *   GET /chats/{chat-id}/messages?$top=50&$orderby=lastModifiedDateTime desc
 *
 * Application permission: ChatMessage.Read.Chat -- the RSC grant made when
 * somebody installed Knowva into this chat. Graph checks it per chat id, so
 * this 403s for any chat Knowva was not added to.
 *
 * NO SERVER-SIDE SEARCH. This endpoint supports only $top, $orderby, and a
 * date-range $filter on createdDateTime / lastModifiedDateTime -- there is no
 * $search and no keyword filter. So "search" here means: pull the recent
 * window, match keywords in this process. That is the single biggest constraint
 * shaping search_conversations, and the reason the scale ceiling in that file
 * is where it is.
 *
 * (Microsoft Search -- POST /search/query with entityType chatMessage -- *does*
 * do server-side keyword search over chat messages. It is rejected here on
 * purpose: it is delegated-only and searches every chat the user belongs to,
 * including ones Knowva was never installed in. That would quietly break
 * "installation is the consent" -- Knowva would be reading conversations nobody
 * added it to. Worse scaling, correct boundary.)
 *
 * Docs:
 *   https://learn.microsoft.com/en-us/graph/api/chat-list-messages?view=graph-rest-1.0
 */

/** One human-authored chat message, flattened to what attribution needs. */
export interface ChatMessageRecord {
  id: string;
  chatId: string;
  authorName: string;
  authorId?: string;
  /** ISO 8601, UTC, straight from Graph. */
  createdDateTime: string;
  /** Body converted to plain text. */
  text: string;
}

interface GraphIdentity {
  id?: string;
  displayName?: string;
}

interface GraphChatMessage {
  id?: string;
  chatId?: string;
  messageType?: string;
  createdDateTime?: string;
  deletedDateTime?: string | null;
  from?: {
    user?: GraphIdentity | null;
    application?: GraphIdentity | null;
    device?: GraphIdentity | null;
  } | null;
  body?: { contentType?: string; content?: string };
}

interface MessageListResponse {
  value?: GraphChatMessage[];
  "@odata.nextLink"?: string;
}

// $top's documented maximum on this endpoint.
const PAGE_SIZE = 50;

/**
 * Pages to pull per chat. 2 x 50 = the ~100 most recent messages. Deliberately
 * shallow: with no server-side search, every extra page is a full round trip
 * that mostly returns messages no keyword will match.
 */
const MAX_PAGES = 2;

export async function listRecentChatMessages(
  graph: GraphClient,
  chatId: string,
  log: ILogger
): Promise<ChatMessageRecord[]> {
  let path =
    `/chats/${encodeURIComponent(chatId)}/messages` +
    `?$top=${PAGE_SIZE}&$orderby=${encodeURIComponent("lastModifiedDateTime desc")}`;

  const raw: GraphChatMessage[] = [];
  for (let page = 1; page <= MAX_PAGES && path; page++) {
    const response = await graph.http.get<MessageListResponse>(path);
    raw.push(...(response.data.value ?? []));

    const next = response.data["@odata.nextLink"];
    if (!next) break;
    path = next;
  }

  const records = raw
    .map((message) => toRecord(message, chatId))
    .filter((record): record is ChatMessageRecord => record !== null);

  log.info(
    `chat-messages: chat ${chatId} -- ${raw.length} activity item(s) fetched, ` +
      `${records.length} usable message(s) after filtering`
  );

  return records;
}

/**
 * Keeps only real, human-authored, still-present messages. Returns null for
 * everything else, each for a distinct reason:
 *
 *  - messageType !== "message" is Teams system noise: "X added Knowva to the
 *    chat", renames, calls started. Graph marks these `systemEventMessage`
 *    (older tenants may send `unknownFutureValue`), gives them a null `from`
 *    and a literal "<systemEventMessage/>" body. They carry no content anyone
 *    would search for, and the *install* event is itself one of them -- so
 *    without this filter Knowva's own arrival becomes a searchable "message".
 *  - deletedDateTime set means somebody deleted it. Retrieving it anyway would
 *    make Knowva a way to read messages that were taken back.
 *  - from.application (and no from.user) is a bot, Knowva's own replies among
 *    them. Excluded so Knowva can never cite itself as a source and launder its
 *    own earlier guess into a colleague's statement of fact.
 */
function toRecord(message: GraphChatMessage, chatId: string): ChatMessageRecord | null {
  if (message.messageType !== "message") return null;
  if (message.deletedDateTime) return null;

  const user = message.from?.user;
  if (!user) return null;

  const text = bodyToText(message.body?.content ?? "", message.body?.contentType);
  if (!text) return null;

  if (!message.id || !message.createdDateTime) return null;

  return {
    id: message.id,
    chatId: message.chatId || chatId,
    // Graph can return a null displayName for a member it can't resolve. Say so
    // rather than dropping the message or inventing a name.
    authorName: user.displayName?.trim() || "an unidentified participant",
    authorId: user.id,
    createdDateTime: message.createdDateTime,
    text,
  };
}

/**
 * Chat bodies come back as either plain text or Teams-flavoured HTML. Strip to
 * text so keyword matching and the model see the same thing a person reads.
 */
function bodyToText(content: string, contentType?: string): string {
  if (!content) return "";
  if (contentType !== "html") return content.trim();

  return content
    // Drop entire non-content elements rather than just their tags, so script
    // and style bodies don't survive as visible text.
    .replace(/<(script|style)\b[^>]*>[\s\S]*?<\/\1>/gi, " ")
    // Block-ish boundaries become spaces so words don't fuse across them.
    .replace(/<\/(p|div|li|tr|h[1-6])>/gi, " ")
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}
