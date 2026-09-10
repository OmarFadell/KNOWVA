import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import { appGraphClient, AppCredentialsUnavailableError } from "../graph/app-client";
import { listCandidateChats, type CandidateChat } from "../graph/installed-chats";
import { listRecentChatMessages, type ChatMessageRecord } from "../graph/chat-messages";
import { chatMessageDeepLink } from "../teams/deep-link";

/**
 * The `search_conversations` tool: finds what colleagues said, in the group
 * chats Knowva has been installed into.
 *
 * Two identities, on purpose (see installed-chats.ts and app-client.ts):
 *   1. The asking user's delegated token lists candidate chats, so results can
 *      only include conversations that user is already a member of.
 *   2. The bot's app-only token, scoped by the per-chat RSC grant, reads the
 *      messages. It can only open chats somebody installed Knowva into.
 * A message has to clear both gates to come back. Neither alone is sufficient.
 *
 * Step 2 is also the installation test. There is no separate "is Knowva
 * installed here?" call: we attempt the read, and a 403 means we cannot read it
 * -- whether because the app was never added, or because its RSC grant is
 * missing. installed-chats.ts explains why that is the right question to ask.
 *
 * ---------------------------------------------------------------------------
 * SCALE CEILING -- READ BEFORE BUILDING ON THIS
 *
 * This is the straightforward implementation, and it does not scale. Every
 * question costs:
 *
 *     1 request to list chats  +  (up to 2 requests x N candidate chats)
 *
 * so ~25 Graph round trips at the current 12-chat cap, all sequential and all
 * inside one Teams turn that has a few seconds before the user gives up. Since
 * installation is now decided by attempting the read, that cost is paid for
 * every candidate chat, including the ones Knowva turns out not to be in.
 * Worse, the work is wasted: the chat-messages endpoint has no $search, so we drag
 * back the last ~100 messages of every chat and throw away the ~99% that do not
 * match. Recall is bounded too -- anything older than that window is invisible
 * regardless of how well it matches. MAX_CHATS_TO_SEARCH below caps the damage
 * and makes the truncation visible to the user rather than silent -- and because
 * the cap now applies to chats that have NOT been pre-filtered by installation,
 * it is load-bearing for recall: a user in 40 group chats with Knowva in only
 * the 20th-most-recent will not find it. Recency ordering in listCandidateChats
 * is what keeps that from being the common case.
 *
 * The replacement is an index, not a bigger cap. Subscribe to Graph change
 * notifications for chat messages (the same RSC permission covers them) and
 * write each message into a searchable store as it arrives, keyed by chat id,
 * with the deep-link fields already resolved. Then a question is one query
 * against that store instead of N round trips, recall stops being a function of
 * how recently something was said, and deletions can be honoured by processing
 * the delete notifications. That store also becomes the natural place to
 * enforce retention, which this version has no way to offer.
 * ---------------------------------------------------------------------------
 */

// Chats searched per question. Twelve keeps a turn inside a few seconds; past
// that a user is watching a typing indicator wondering if the bot died.
const MAX_CHATS_TO_SEARCH = 12;

// Messages handed to the model. Enough to answer from, small enough to leave
// the model room to actually write the answer.
const MAX_RESULTS = 12;

// Tokens shorter than this match too much to be worth scoring on.
const MIN_TERM_LENGTH = 3;

/**
 * Words too common to discriminate between messages. Not a general stopword
 * list -- just enough that "what did we decide about the launch date" scores on
 * "decide", "launch" and "date" rather than on "what" and "about".
 */
const STOPWORDS = new Set([
  "the", "and", "for", "was", "are", "were", "you", "your", "our", "who", "what",
  "when", "where", "why", "how", "did", "does", "is", "it", "its", "this",
  "that", "these", "those", "with", "about", "from", "have", "has", "had", "not",
  "but", "any", "all", "can", "could", "would", "should", "will", "there", "here",
  "they", "them", "their", "say", "said", "says", "get", "got",
]);

export function createConversationSearchTool(
  userGraph: GraphClient,
  tenantId: string | undefined,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "search_conversations",
      description:
        "Search recent messages in the Teams group chats Knowva has been installed in. " +
        "Use this when the user asks what someone said, what was decided or discussed, or " +
        "who mentioned something. Returns individual messages with the author's name, the " +
        "time they sent it, and a link to the message. It searches only group chats Knowva " +
        "was added to AND that the person asking is a member of; it cannot see private " +
        "one-to-one chats, and it only covers roughly the most recent messages in each chat.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "The keywords or topic to look for, e.g. 'launch date decision' or " +
              "'budget approval'. Distinctive words work best; this matches words in " +
              "message text rather than understanding a question.",
          },
        },
        required: ["query"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      const query = (typeof input.query === "string" ? input.query : "").trim();
      if (!query) {
        return {
          content:
            "No search terms were provided. Call search_conversations again with the keywords to look for.",
          isError: true,
        };
      }

      const terms = tokenize(query);
      if (terms.length === 0) {
        return {
          content:
            `"${query}" has no distinctive words to search on. Ask the user for a more specific ` +
            "term -- a project name, a person, a decision -- and try again.",
          isError: true,
        };
      }

      // --- Gate 1: which chats may this user see at all? -------------------
      let chats: CandidateChat[];
      try {
        chats = await listCandidateChats(userGraph, log);
      } catch (err) {
        return classifyEnumerationError(err, log);
      }

      if (chats.length === 0) {
        // Not an error: this user is in no group chats at all. Distinct from
        // "in group chats, but Knowva isn't in any of them", handled below.
        log.info("conversation-search: user is in no group or meeting chats");
        return {
          content:
            "This user isn't a member of any Teams group chat, so there are no conversations to " +
            "search. Tell them plainly, and do NOT answer the question from your own knowledge " +
            "as though you had searched.",
        };
      }

      const searched = chats.slice(0, MAX_CHATS_TO_SEARCH);
      const skipped = chats.length - searched.length;

      // --- Gate 2: read those chats with the RSC-scoped app token ----------
      let appGraph: GraphClient;
      try {
        appGraph = await appGraphClient();
      } catch (err) {
        log.error(
          "conversation-search: could not get an app-only Graph token",
          describeError(err),
          err
        );
        if (err instanceof AppCredentialsUnavailableError) {
          return {
            content:
              "Conversation search isn't configured on this deployment -- Knowva has no application " +
              "credentials to read chat messages with. Tell the user it's a setup problem on our side.",
            isError: true,
          };
        }
        return {
          content:
            "Knowva couldn't authenticate to read chat messages. Tell the user it's a problem on " +
            "our side, not theirs.",
          isError: true,
        };
      }

      const messages: ChatMessageRecord[] = [];
      const chatById = new Map<string, CandidateChat>();
      let blockedChats = 0;
      let failedChats = 0;

      // Sequential rather than parallel: Graph throttles per-app, and a burst of
      // a dozen simultaneous message reads is the shape that earns a 429.
      for (const chat of searched) {
        chatById.set(chat.id, chat);
        try {
          messages.push(...(await listRecentChatMessages(appGraph, chat.id, log)));
        } catch (err) {
          const status = statusOf(err);
          if (status === 403 || status === 401) {
            // THIS IS THE INSTALLATION CHECK. A 403 here is the ordinary,
            // expected answer for a chat Knowva was never added to -- which is
            // most of them -- so it is logged at debug, not warn. It also covers
            // "installed, but the RSC grant is missing or blocked by policy";
            // those are indistinguishable and equally unreadable, so they are
            // treated identically and reported honestly below.
            blockedChats++;
            log.debug(
              `conversation-search: chat ${chat.id} not readable (status ${status}) -- ` +
                "Knowva is probably not installed there; skipping"
            );
          } else {
            failedChats++;
            log.error(
              `conversation-search: failed to read chat ${chat.id}`,
              describeError(err),
              err
            );
          }
        }
      }

      // Not one chat was readable. Overwhelmingly the common cause is simply
      // that Knowva has not been added to any of them, so that is named first --
      // but the RSC possibilities are named too, because they look identical
      // from here and telling the user only the likely cause would be guessing.
      if (blockedChats > 0 && blockedChats === searched.length) {
        return {
          content:
            `Knowva could not read any of the ${searched.length} group chat(s) it tried. The usual ` +
            "reason is simply that it hasn't been added to them -- it can only read a chat somebody " +
            "installed it into. Tell the user to add Knowva to the group chat they care about. If " +
            "it IS already installed there, then either the app was added before it asked for " +
            "message access (remove and re-add it), or resource-specific consent is blocked by " +
            "tenant policy (an administrator would need to check). Do NOT answer the question from " +
            "your own knowledge.",
        };
      }

      if (failedChats > 0 && failedChats + blockedChats === searched.length) {
        return {
          content:
            "Conversation search failed against every chat it tried. Tell the user it's a problem " +
            "on our side, not theirs, and to try again shortly.",
          isError: true,
        };
      }

      // --- Rank in-process, because Graph cannot rank for us ---------------
      const ranked = messages
        .map((message) => ({ message, score: score(message.text, terms) }))
        .filter((scored) => scored.score > 0)
        .sort(
          (a, b) =>
            b.score - a.score ||
            // Tie-break on recency: for "what did we decide", the later message
            // is almost always the one that matters.
            Date.parse(b.message.createdDateTime) - Date.parse(a.message.createdDateTime)
        )
        .slice(0, MAX_RESULTS);

      log.info(
        `conversation-search: "${query}" -- ${searched.length} chat(s) searched, ` +
          `${messages.length} message(s) considered, ${ranked.length} matched ` +
          `(blocked=${blockedChats} failed=${failedChats} skipped=${skipped})`
      );

      if (ranked.length === 0) {
        // Chats were readable; nothing in them matched. Distinct from both
        // "not installed anywhere" and "access denied", and the model must say so.
        return {
          content:
            `No messages matching "${query}" were found in the group chat(s) Knowva can read ` +
            `(it tried ${searched.length}). Tell the user nothing came up, mention that only recent ` +
            "messages are searchable, and suggest more specific terms. Do NOT answer from your own " +
            "knowledge as though you had found something.",
        };
      }

      return {
        content: formatResults(
          ranked.map((scored) => scored.message),
          chatById,
          tenantId,
          { searchedChats: searched.length, skipped, blockedChats }
        ),
      };
    },
  };
}

function tokenize(query: string): string[] {
  return Array.from(
    new Set(
      query
        .toLowerCase()
        .split(/[^a-z0-9]+/)
        .filter((term) => term.length >= MIN_TERM_LENGTH && !STOPWORDS.has(term))
    )
  );
}

/**
 * How well one message matches the search terms.
 *
 * Deliberately simple: count distinct terms present, with a small bonus for a
 * term appearing more than once. It is a keyword filter, not relevance ranking
 * -- the model does the actual judging once it has the candidates, and the tool
 * description says so rather than implying semantic search.
 */
function score(text: string, terms: string[]): number {
  const haystack = text.toLowerCase();
  let total = 0;

  for (const term of terms) {
    let occurrences = 0;
    let index = haystack.indexOf(term);
    while (index !== -1) {
      occurrences++;
      index = haystack.indexOf(term, index + term.length);
    }
    if (occurrences > 0) total += 1 + Math.min(occurrences - 1, 2) * 0.25;
  }

  return total;
}

/**
 * Renders matches for the model. Every message carries author, UTC timestamp
 * and a deep link, because the system prompt forbids citing a colleague without
 * all three -- so the tool must never return a message missing any of them.
 */
function formatResults(
  messages: ChatMessageRecord[],
  chatById: Map<string, CandidateChat>,
  tenantId: string | undefined,
  stats: { searchedChats: number; skipped: number; blockedChats: number }
): string {
  const blocks = messages.map((message, i) => {
    const chat = chatById.get(message.chatId);
    const chatName = chat?.topic?.trim() || "an unnamed group chat";
    const link = chatMessageDeepLink(message.chatId, message.id, chat?.tenantId || tenantId);

    return (
      `[${i + 1}] ${message.authorName} in "${chatName}" at ${message.createdDateTime} (UTC)\n` +
      `    link: ${link}\n` +
      `    said: ${message.text}`
    );
  });

  const caveats: string[] = [
    `Searched the ${stats.searchedChats} group chat(s) Knowva is installed in, ` +
      "covering only their recent messages.",
  ];
  if (stats.skipped > 0) {
    caveats.push(
      `${stats.skipped} further installed chat(s) were NOT searched (per-question limit). ` +
        "Mention this if the user seems to expect complete coverage."
    );
  }
  if (stats.blockedChats > 0) {
    caveats.push(
      `${stats.blockedChats} chat(s) refused access and were skipped, so results may be incomplete.`
    );
  }

  return (
    "Conversation search results. These are things real colleagues actually wrote. " +
    "Attribute every one of them: name the person, give the time, and link the message. " +
    "Never restate what someone said as if it were established fact.\n\n" +
    blocks.join("\n\n") +
    "\n\n" +
    caveats.join(" ")
  );
}

function statusOf(err: unknown): number | undefined {
  const e = err as { statusCode?: number; response?: { status?: number } };
  return typeof e?.statusCode === "number" ? e.statusCode : e?.response?.status;
}

/**
 * Failures listing chats. Kept separate from the per-chat message failures
 * above because the causes are different: this one is about the *user's*
 * delegated Chat.ReadBasic scope, not about RSC.
 */
function classifyEnumerationError(err: unknown, log: ILogger): AgentToolResult {
  log.error("conversation-search: could not list installed chats", describeError(err), err);

  const status = statusOf(err);

  if (status === 401) {
    return {
      content:
        "Conversation search could not authenticate as the user. Tell them to sign out and sign " +
        "back in, then retry.",
      isError: true,
    };
  }

  if (status === 403) {
    return {
      content:
        "Conversation search was denied when listing chats. Knowva most likely hasn't been granted " +
        "the Chat.ReadBasic permission yet, which it needs to tell which chats it's installed in. " +
        "Tell the user it isn't set up yet and an administrator needs to grant consent.",
      isError: true,
    };
  }

  if (status === 429) {
    return {
      content:
        "Conversation search is temporarily rate-limited. Tell the user to try again in a minute.",
      isError: true,
    };
  }

  return {
    content:
      `Conversation search failed while listing chats${status ? ` (status ${status})` : ""}. ` +
      "Tell the user it's a problem on our side, not theirs.",
    isError: true,
  };
}
