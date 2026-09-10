import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";

/**
 * Lists the group chats the asking user is in, as *candidates* for a
 * conversation search.
 *
 *   GET /me/chats?$top=50&$orderby=lastMessagePreview/createdDateTime desc
 *
 * The `graph` client must carry the *asking user's* delegated token. That is
 * deliberate and does two jobs:
 *
 *  1. Least privilege. This needs only Chat.ReadBasic, which returns chat
 *     metadata (id, topic, type) and never message content. The application
 *     alternatives -- Chat.ReadBasic.All / Chat.Read.All -- are tenant-wide
 *     admin-consented grants over *every* chat in the org, which is exactly the
 *     blast radius RSC exists to avoid.
 *  2. Privacy, structurally. Because it runs as the asker, the result can only
 *     ever contain chats that user is already a member of. Knowva physically
 *     cannot surface a chat the asker isn't in, even one it is installed in.
 *
 * ---------------------------------------------------------------------------
 * THIS FUNCTION DOES NOT DETERMINE WHETHER KNOWVA IS INSTALLED. Read on before
 * "optimising" it back into something that does.
 *
 * It used to try, with:
 *
 *   GET /me/chats?$filter=installedApps/any(a:a/teamsApp/id eq '{appId}')
 *
 * That query is NOT supported and Graph answers it with a 404 -- not the 400
 * you would expect from a bad filter, which is what made it read like a routing
 * or permissions problem rather than a malformed query. It is documented as
 * Example 5 on the chat-list reference page; the documentation is wrong, or at
 * least not true of /me/chats. Do not restore it.
 *
 * It was also wrong on a second, independent axis: it matched `teamsApp/id`
 * against TEAMS_APP_ID. Those are different identifiers. Per the teamsApp
 * resource reference, `id` is "the app ID generated for the catalog [...]
 * different from the developer-provided ID found within the Teams zip app
 * package", while `externalId` is the developer-provided one. TEAMS_APP_ID is
 * what appPackage/manifest.json interpolates into its own `id` field, so it is
 * the *externalId*. Matching it against `teamsApp/id` returns zero rows even
 * when the query shape is valid. Any future code that matches this app against
 * an installation must compare TEAMS_APP_ID to `teamsApp/externalId`.
 *
 * Installation is instead determined by the caller, by attempting the message
 * read and treating 403 as "not installed here" -- see the note below and
 * src/tools/conversation-search.ts.
 * ---------------------------------------------------------------------------
 *
 * WHY 403-AS-NOT-INSTALLED, RATHER THAN CHECKING /chats/{id}/installedApps
 *
 * The documented way to ask "is this app installed in this chat?" is:
 *
 *   GET /chats/{id}/installedApps?$expand=teamsApp&$filter=teamsApp/externalId eq '{id}'
 *
 * That works and is well documented. It was rejected here for three reasons:
 *
 *  1. It needs a delegated permission we do not hold and would have to add:
 *     TeamsAppInstallation.ReadForChat. That means another manifest change,
 *     reprovision and re-consent cycle to answer a question we can already
 *     answer for free.
 *  2. It costs one extra round trip per candidate chat, *on top of* the message
 *     read we were going to make anyway. Attempting the read is strictly fewer
 *     requests, and this path is already round-trip-bound.
 *  3. It answers the wrong question. "Is the app installed?" is a proxy; what
 *     we actually need is "can Knowva read this chat's messages?". Those come
 *     apart in a case we have already hit: a chat where Knowva was installed
 *     *before* ChatMessage.Read.Chat was added to the manifest reports as
 *     installed, but the RSC grant was never made and the read still 403s.
 *     Attempting the read cannot be fooled that way.
 *
 * The cost of this choice: a 403 no longer distinguishes "not installed" from
 * "installed but RSC blocked by tenant policy". Both mean "we cannot read it",
 * so the behaviour is identical -- but the message shown to the user has to
 * name both possibilities honestly. conversation-search.ts does that.
 */

/** Graph `chat`, narrowed to the fields this feature uses. */
export interface CandidateChat {
  id: string;
  topic?: string;
  chatType?: string;
  webUrl?: string;
  tenantId?: string;
}

interface ChatListResponse {
  value?: CandidateChat[];
  "@odata.nextLink"?: string;
}

// $top's documented maximum on this endpoint.
const PAGE_SIZE = 50;

// Hard stop on paging. See the SCALE note in src/tools/conversation-search.ts:
// past a few dozen chats this whole approach needs replacing, so following
// pages forever would only make a slow answer slower.
const MAX_PAGES = 4;

/**
 * Most-recently-active first. This matters more than it looks: without the
 * server-side install filter, the caller's per-question chat cap now applies to
 * *unfiltered* chats, so which chats end up at the front of this list decides
 * what gets searched at all. Recent activity is the best cheap proxy for "the
 * chat the user means".
 */
const ORDER_BY = "lastMessagePreview/createdDateTime desc";

/**
 * Chat types worth searching. Group chats and meeting chats are multi-person
 * conversations somebody could have added Knowva to. `oneOnOne` is excluded:
 * the app is installed in personal scope for every user who ever opened it, so
 * without this filter a search would trawl each user's private 1:1 thread with
 * the bot -- which is not what "search our conversations" means to anyone.
 *
 * Applied client-side. The previous version expressed the same intent through a
 * server-side $filter that turned out not to be supported; doing it here cannot
 * fail, at the cost of fetching a few rows we discard.
 */
const SEARCHABLE_CHAT_TYPES = new Set(["group", "meeting"]);

export async function listCandidateChats(
  graph: GraphClient,
  log: ILogger
): Promise<CandidateChat[]> {
  // $orderby on this endpoint is documented, but so was the $filter that turned
  // out to 404 -- so it is treated as best-effort rather than trusted. If the
  // ordered request fails we retry once unordered, because unordered results are
  // far better than no results.
  let chats: CandidateChat[];
  try {
    chats = await fetchChats(
      graph,
      `/me/chats?$top=${PAGE_SIZE}&$orderby=${encodeURIComponent(ORDER_BY)}`,
      log
    );
  } catch (err) {
    log.warn(
      "installed-chats: ordered chat listing failed, retrying without $orderby " +
        "(results will not be recency-ordered)",
      { message: (err as { message?: string })?.message }
    );
    chats = await fetchChats(graph, `/me/chats?$top=${PAGE_SIZE}`, log);
  }

  const searchable = chats.filter(
    (chat) => chat.id && SEARCHABLE_CHAT_TYPES.has(chat.chatType ?? "")
  );

  log.info(
    `installed-chats: ${chats.length} chat(s) visible to this user, ` +
      `${searchable.length} searchable candidate(s) (group/meeting). ` +
      "Whether Knowva is installed in each is determined by the message read."
  );

  return searchable;
}

async function fetchChats(
  graph: GraphClient,
  firstPage: string,
  log: ILogger
): Promise<CandidateChat[]> {
  const chats: CandidateChat[] = [];
  let path = firstPage;

  for (let page = 1; page <= MAX_PAGES && path; page++) {
    const response = await graph.http.get<ChatListResponse>(path);
    chats.push(...(response.data.value ?? []));

    const next = response.data["@odata.nextLink"];
    if (!next) break;
    if (page === MAX_PAGES) {
      log.warn(
        `installed-chats: stopped at the ${MAX_PAGES}-page cap; some chats were not considered`
      );
      break;
    }
    path = next;
  }

  return chats;
}
