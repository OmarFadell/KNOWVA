import type { ILogger } from "@microsoft/teams.common";
import type { AgentTool, AgentToolResult } from "../agent/loop";
import type { ActingIdentityResolver } from "../auth/acting-identity";
import {
  dropAtlassianTokens,
  getValidAccessToken,
  isAtlassianConfigured,
  missingAtlassianScopes,
  type AtlassianTokenFailure,
} from "../auth/atlassian";
import {
  getConfluencePage,
  searchConfluence,
  type ConfluenceFailure,
  type ConfluenceSearchHit,
} from "../confluence/client";

/**
 * The Confluence tools: `search_confluence` and `get_confluence_page`.
 *
 * ===========================================================================
 * THIS IS THE ONE ATLASSIAN TOOL WITH GENUINE PER-USER SCOPING.
 *
 * Every other Atlassian-facing thing Knowva does is an approximation of
 * per-user access, because Xray Cloud offers no per-user auth at all --
 * find_my_xray_items searches on a shared key and filters to items naming the
 * asker. These two tools are different in kind: they call Confluence with the
 * asking user's OWN OAuth token, and Confluence enforces their real space and
 * page permissions server-side. A page they cannot read is never returned.
 *
 * So there is no space allowlist and no post-filtering, and adding either would
 * be a mistake rather than a hardening: Confluence has already decided. The
 * search covers every space the signed-in user can access, which is exactly the
 * set they would see searching Confluence themselves.
 *
 * THE CONFUSED-DEPUTY STAKES ARE THE SAME AS GITHUB'S, THOUGH. Knowva holds a
 * live Confluence token per user, and picking the wrong one is undetectable
 * from Confluence's side -- it receives a valid bearer token and answers for
 * whoever it belongs to. So the acting identity is resolved from the Graph
 * token and cross-checked against the id Teams put on the activity, exactly as
 * src/tools/github-access.ts does, and both tools refuse on any doubt. In a
 * group chat that is the difference that matters: the turn runs on the asker's
 * token, so the asker's Confluence session is the only one in play.
 * ===========================================================================
 *
 * TWO STAGES, MATCHING SHAREPOINT AND EMAIL. search_confluence returns titles,
 * spaces, timestamps and short excerpts; get_confluence_page pulls one full
 * body on demand. That keeps a broad search affordable, and means a page nobody
 * needed is never read at all -- a privacy property as much as a cost one.
 *
 * BOTH TOOLS SHARE ONE TOKEN GATE (resolveConfluenceAccess below), for the same
 * reason src/tools/github-access.ts exists: two tools answering "may this user
 * do this?" from two copies of the logic will eventually answer it differently,
 * and the one that drifts will be the one nobody reads.
 */

/** Search hits per query. Confluence excerpts are short; this keeps a turn affordable. */
const RESULTS_PER_SEARCH = 10;

/** Longest query accepted. */
const MAX_QUERY_LENGTH = 300;

interface ConfluenceAccess {
  ok: boolean;
  token?: string;
  cloudId?: string;
  siteUrl?: string;
  /** The Entra object id this access was resolved for, so a dead token can be dropped. */
  userId?: string;
  /** Set when not ok: a ready-made result for the tool to return as-is. */
  refusal?: AgentToolResult;
}

/**
 * The gate both tools pass through: resolve who is asking, then their live
 * Confluence token.
 *
 * Returns a ready-made `refusal` rather than a reason code, because both
 * callers would otherwise translate the same failures into the same messages.
 */
async function resolveConfluenceAccess(
  identity: ActingIdentityResolver,
  log: ILogger
): Promise<ConfluenceAccess> {
  if (!isAtlassianConfigured()) return { ok: false, refusal: notConfiguredResult() };

  // Nothing that spends a credential may run before this.
  const acting = await identity.resolve();
  if (!acting.ok) {
    log.warn(`confluence: refusing -- acting identity ${acting.reason}`);
    return { ok: false, refusal: identityRefusalResult(acting.reason) };
  }

  const userId = acting.userId as string;

  const token = await getValidAccessToken(userId, log);
  if (!token.ok) return { ok: false, refusal: needsAuthResult(token.reason) };

  // PRE-FLIGHT SCOPE CHECK. Atlassian fixes a token's scopes at consent time,
  // so a session created before the Confluence scopes were added is valid,
  // refreshable, and completely unable to call Confluence -- it would 401 on
  // every request, forever. Catching it here saves a pointless round trip and,
  // more importantly, lets the user be told the actual reason instead of
  // being shown a generic auth failure that looks like a revocation.
  const missing = missingAtlassianScopes(userId);
  if (missing.length > 0) {
    log.warn(
      `confluence: ${userId} has a session missing ${missing.join(", ")}; forcing re-consent ` +
        "rather than making a call that can only 401"
    );
    dropAtlassianTokens(userId);
    return { ok: false, refusal: scopeMismatchResult(missing) };
  }

  // The cloud id was resolved once at sign-in and cached on the identity. One
  // cloud id addresses both /ex/jira/ and /ex/confluence/ for the same site, so
  // there is nothing to look up here. Its absence would mean a session stored
  // before that caching existed.
  if (!token.cloudId) {
    log.error("confluence: session has no cloudId; cannot address the Confluence API");
    return { ok: false, refusal: noCloudIdResult() };
  }

  return {
    ok: true,
    token: token.token,
    cloudId: token.cloudId,
    siteUrl: token.siteUrl,
    userId,
  };
}

export function createConfluenceSearchTool(
  identity: ActingIdentityResolver,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "search_confluence",
      description:
        "Search Confluence pages. Only pages the person you are talking to can already see are " +
        "searched -- Confluence applies their own permissions, so this can never reveal a page " +
        "they lack access to. Use it when the user asks about documentation, specs, runbooks, " +
        "meeting notes, or anything that would live in Confluence. Returns page titles, the " +
        "space each lives in, when it was last updated, a short excerpt and a link. Call " +
        "get_confluence_page afterwards to read one in full.",
      inputSchema: {
        type: "object",
        properties: {
          query: {
            type: "string",
            description:
              "The keywords to search for. Plain words, not query syntax -- e.g. 'VPN setup " +
              `onboarding'. At most ${MAX_QUERY_LENGTH} characters.`,
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
            "No search terms were provided. Call search_confluence again with the keywords to " +
            "look for.",
          isError: true,
        };
      }
      if (query.length > MAX_QUERY_LENGTH) {
        return {
          content:
            `That query is ${query.length} characters; the limit is ${MAX_QUERY_LENGTH}. ` +
            "Shorten it to the distinctive terms and try again.",
          isError: true,
        };
      }

      const access = await resolveConfluenceAccess(identity, log);
      if (!access.ok) return access.refusal as AgentToolResult;

      const result = await searchConfluence(
        access.token as string,
        access.cloudId as string,
        access.siteUrl,
        query,
        RESULTS_PER_SEARCH,
        log
      );

      if (!result.ok) {
        // A token Confluence refuses cannot be made to work by retrying, so it
        // is discarded here rather than re-presented on every later question.
        // The identity half survives, so find_my_xray_items is unaffected.
        if (result.reason === "unauthorized" || result.reason === "scope-mismatch") {
          dropAtlassianTokens(access.userId as string);
        }
        return confluenceFailureResult(result.reason, result.detail, "that search");
      }

      const hits = result.hits || [];
      if (hits.length === 0) {
        return {
          // NOT isError. Finding nothing is a real answer.
          content:
            `No Confluence pages matched "${query}". Tell the user nothing was found, and ` +
            "remind them this only covers spaces their own Confluence account can see. Do NOT " +
            "answer from your own knowledge as though you had found something.",
        };
      }

      return { content: formatSearchResults(hits, query, result.total) };
    },
  };
}

export function createConfluencePageTool(
  identity: ActingIdentityResolver,
  log: ILogger
): AgentTool {
  return {
    definition: {
      name: "get_confluence_page",
      description:
        "Read one Confluence page in full, by its page id. Use the id exactly as " +
        "search_confluence returned it. Search first to find the page, then call this for the " +
        "one you actually need. Returns the page body as text, plus its title, space, and when " +
        "it was last updated and by whom.",
      inputSchema: {
        type: "object",
        properties: {
          pageId: {
            type: "string",
            description:
              "The Confluence page id, exactly as search_confluence returned it. A number like " +
              "'123456789'. Do not invent one or guess it from a URL.",
          },
        },
        required: ["pageId"],
        additionalProperties: false,
      },
    },

    async run(input: Record<string, unknown>): Promise<AgentToolResult> {
      const pageId = (typeof input.pageId === "string" ? input.pageId : "").trim();
      if (!pageId) {
        return {
          content:
            "No page id was provided. Call search_confluence first, then pass the id from one " +
            "of its results.",
          isError: true,
        };
      }
      // Confluence content ids are numeric. Rejecting anything else here turns
      // a hallucinated id into a clear instruction rather than a 404 the model
      // might read as "the page does not exist".
      if (!/^\d+$/.test(pageId)) {
        return {
          content:
            `"${pageId}" is not a Confluence page id. Ids are numeric and come from ` +
            "search_confluence results -- search first and use the id it gives you, rather than " +
            "taking one from a URL or inventing it.",
          isError: true,
        };
      }

      const access = await resolveConfluenceAccess(identity, log);
      if (!access.ok) return access.refusal as AgentToolResult;

      const result = await getConfluencePage(
        access.token as string,
        access.cloudId as string,
        access.siteUrl,
        pageId,
        log
      );

      if (!result.ok) {
        // 404 and 403 are worth separating here in a way they are not for
        // search. Confluence returns 404 for a page that exists but which this
        // user cannot see, so the honest answer is "you may not have access",
        // not "it does not exist".
        if (result.reason === "not-found") {
          return {
            content:
              `No Confluence page with id ${pageId} came back. Either it does not exist, or it ` +
              "exists and this user cannot see it -- Confluence does not distinguish the two, " +
              "and neither should you. Say it could not be opened and that they may not have " +
              "access, rather than asserting it does not exist.",
            isError: true,
          };
        }
        if (result.reason === "unauthorized" || result.reason === "scope-mismatch") {
          dropAtlassianTokens(access.userId as string);
        }
        return confluenceFailureResult(result.reason, result.detail, "that page");
      }

      const page = result.page as NonNullable<typeof result.page>;
      return { content: formatPage(page) };
    },
  };
}

/**
 * Wraps search hits with the citation rules.
 *
 * The rules are as strict as the ones for SharePoint documents, and for the
 * same reason: an excerpt is a fragment, and a model given fragments will
 * cheerfully describe the whole page it came from.
 */
function formatSearchResults(
  hits: ConfluenceSearchHit[],
  query: string,
  total: number | undefined
): string {
  const lines = hits.map((hit) => {
    const parts = [
      `- "${hit.title}"`,
      hit.spaceName ? ` in the ${hit.spaceName} space` : hit.spaceKey ? ` in ${hit.spaceKey}` : "",
      hit.lastModified ? ` — last updated ${hit.lastModified}` : "",
      hit.lastModifiedBy ? ` by ${hit.lastModifiedBy}` : "",
      `\n  page id: ${hit.id}`,
      hit.url ? `\n  link: ${hit.url}` : "",
      hit.excerpt ? `\n  excerpt: ${hit.excerpt}` : "",
    ];
    return parts.join("");
  });

  const more =
    typeof total === "number" && total > hits.length
      ? ` Confluence reports ${total} matches in total; these are the first ${hits.length}.`
      : "";

  return (
    `Confluence pages matching "${query}", searched as this user with their own Confluence ` +
    `permissions.${more} These cover only spaces they can already see.\n\n` +
    lines.join("\n\n") +
    "\n\n" +
    "These are EXCERPTS, not pages. Each is a fragment Confluence chose because it contains the " +
    "search terms, and it is not a summary of the page.\n\n" +
    "Cite everything you take from these: give the page title, the space it is in, and link it " +
    "in Markdown using the link above. Timestamps are UTC (ISO 8601) -- present them readably " +
    "and say they are UTC. If an excerpt does not actually answer the question, call " +
    "get_confluence_page with that page's id and read it properly rather than guessing from the " +
    "fragment. Never describe what a page says beyond what appears above."
  );
}

/**
 * Wraps a full page with the paraphrase rule.
 *
 * WHY PARAPHRASE IS THE DEFAULT HERE and not for, say, a code file: a
 * Confluence page is somebody's writing, often long, and reproducing it wholesale
 * into a Teams reply is both useless to read and a quiet way of relocating
 * internal documentation into a chat transcript. The page is one click away via
 * the link; the model's job is to answer the question from it.
 */
function formatPage(page: {
  id: string;
  title: string;
  spaceKey?: string;
  spaceName?: string;
  lastModified?: string;
  lastModifiedBy?: string;
  version?: number;
  text: string;
  truncated: boolean;
  url?: string;
}): string {
  const header = [
    `Confluence page "${page.title}"`,
    page.spaceName ? ` in the ${page.spaceName} space` : page.spaceKey ? ` in ${page.spaceKey}` : "",
    page.lastModified ? `, last updated ${page.lastModified}` : "",
    page.lastModifiedBy ? ` by ${page.lastModifiedBy}` : "",
    page.version ? ` (version ${page.version})` : "",
    ". Read with this user's own Confluence permissions.",
  ].join("");

  return (
    header +
    (page.url ? `\nLink: ${page.url}` : "") +
    `\npage id: ${page.id}` +
    (page.truncated
      ? "\n\nTHIS PAGE WAS TRUNCATED -- you are seeing the beginning only. Say so if you answer " +
        "from it, and do not claim to have read the whole page or infer what the rest contains."
      : "") +
    // v2 reports the last editor as an account id, not a name, so the page
    // result carries no editor. Without this the model fills the gap -- it will
    // attribute the edit to whoever the search result mentioned, or to the
    // person asking. Saying nothing is the correct answer here.
    (page.lastModifiedBy
      ? ""
      : "\n\nYou have NOT been told who last edited this page. Do not name or guess an editor, " +
        "and do not carry one over from a search result.") +
    "\n\nPARAPHRASE THIS, DO NOT REPRODUCE IT. Answer the user's question in your own words and " +
    "quote at most a sentence or two where the exact wording genuinely matters. Do not paste " +
    "large sections, and do not reproduce the page as a whole -- it is long, the user has the " +
    "link, and a chat reply is not the place to relocate internal documentation. Cite the page " +
    "title and space, and link it in Markdown. Timestamps are UTC.\n\n" +
    "--- page content ---\n" +
    page.text
  );
}

// --- Error results -------------------------------------------------------

function notConfiguredResult(): AgentToolResult {
  return {
    content:
      "Confluence isn't set up on this deployment -- Knowva has no Atlassian OAuth app " +
      "configured. Tell the user it's a setup problem on our side, not something they did.",
    isError: true,
  };
}

function identityRefusalResult(reason: string | undefined): AgentToolResult {
  return {
    content:
      reason === "mismatch"
        ? "Knowva could not safely establish whose Confluence account to act as, so it did " +
          "nothing. This is a safety stop. Tell the user something is wrong with their sign-in " +
          "and to sign out and back in."
        : "Knowva could not confirm who it would be acting as in Confluence, so it did nothing. " +
          "This is a safety stop, not a permissions problem. Tell the user to sign out and sign " +
          "back in, then try again.",
    isError: true,
  };
}

function noCloudIdResult(): AgentToolResult {
  return {
    content:
      "Knowva has a Confluence session for this user but no record of which Atlassian site it " +
      "belongs to, so it could not make the request. Tell them to run **/jira-signout** and " +
      "connect again, which will re-establish it.",
    isError: true,
  };
}

/**
 * The "sign in first" result: prose for the model, plus the signal that makes
 * app.ts attach a sign-in card (see AgentToolSignal in src/agent/loop.ts).
 *
 * The reasons are kept apart because they read very differently to somebody who
 * has been using this successfully for weeks. "Connect Confluence" is right for
 * a user who never has; "your connection expired" is right for one whose grant
 * lapsed, and telling them to connect would imply their earlier sign-in never
 * happened.
 */
function needsAuthResult(reason: AtlassianTokenFailure | undefined): AgentToolResult {
  if (reason === "not-configured") return notConfiguredResult();

  const explanation =
    reason === "session-expired"
      ? "This user's Confluence connection has expired and cannot be renewed automatically, so " +
        "nothing was searched."
      : reason === "refresh-failed"
        ? "This user's Confluence connection could not be renewed -- most likely they revoked " +
          "it in their Atlassian account settings. Nothing was searched."
        : "This user hasn't connected their Atlassian account yet, so nothing was searched.";

  return {
    // No URL in here on purpose -- the card carries it.
    content:
      explanation +
      " Knowva is showing them a sign-in button, so do NOT paste a link or invent one. Tell " +
      "them briefly that you need them to connect Confluence first and that there's a button " +
      "just below your message, then stop. Say nothing about what you might have found -- you " +
      "searched nothing. Note this is an Atlassian sign-in, entirely separate from their " +
      "Microsoft Teams sign-in and from GitHub.",
    signal: { kind: "needs-auth", provider: "atlassian" },
  };
}

/**
 * The message for a token that is valid but under-scoped.
 *
 * WORDED CAREFULLY, because the obvious phrasing is wrong in a way that wastes
 * the user's time. Nothing has been revoked, nothing has expired, and they did
 * nothing. Their connection simply predates Confluence being added, and
 * Atlassian will not upgrade an existing grant -- only a fresh consent can.
 * Telling them their access "was revoked" sends them to check an Atlassian
 * settings page where everything looks fine.
 */
function scopeMismatchResult(missing: string[]): AgentToolResult {
  return {
    content:
      "This user's Atlassian connection was made before Knowva could read Confluence, so it " +
      "does not carry Confluence permission and nothing was searched. Nothing is broken and " +
      "nothing was revoked -- Atlassian fixes permissions when you connect, so the existing " +
      "connection simply cannot be upgraded in place. Knowva is showing them a button to " +
      "reconnect, which fixes it in one click. Tell them briefly that they need to reconnect " +
      "Atlassian to give Knowva access to Confluence, that there's a button just below your " +
      "message, and that it is not a problem with their account. Do NOT tell them their access " +
      "was revoked or that they need to change anything in Atlassian. Do not paste or invent a " +
      "link. Say nothing about what you might have found -- you searched nothing." +
      ` (Technical detail, not for the user: missing scopes ${missing.join(", ")}.)`,
    signal: { kind: "needs-auth", provider: "atlassian" },
  };
}

function confluenceFailureResult(
  reason: ConfluenceFailure | undefined,
  detail: string | undefined,
  what: string
): AgentToolResult {
  switch (reason) {
    case "scope-mismatch":
      // Reached when the pre-flight check could not catch it -- a session
      // stored before scopes were recorded, so missingAtlassianScopes() had
      // nothing to compare and correctly declined to guess. Confluence's own
      // 401 body is the fallback signal.
      return scopeMismatchResult([]);

    case "unauthorized":
      // A token the auth layer believed was live, rejected by Confluence with
      // something other than a scope complaint -- so a revoked or genuinely
      // dead grant. Distinct from scope-mismatch above, which is far more
      // common and means the opposite thing about the user's account.
      return {
        content:
          "Confluence rejected this user's credentials. Their connection may have been revoked " +
          "on Atlassian's side. Tell them to reconnect by asking again -- Knowva will offer a " +
          "fresh sign-in button -- and note this is the Atlassian connection, not their Teams " +
          "sign-in.",
        isError: true,
        signal: { kind: "needs-auth", provider: "atlassian" },
      };

    case "forbidden":
      return {
        content:
          `Confluence refused ${what}. This user's account does not have access to it. Tell ` +
          "them plainly and suggest they check with whoever administers that space. Do not " +
          "guess at what the result would have been.",
        isError: true,
      };

    case "bad-query":
      return {
        content:
          `Confluence rejected ${what} as malformed. That is a bug on our side, not the user's ` +
          "problem -- say so plainly rather than suggesting they rephrase." +
          (detail ? ` (Technical detail, not for the user: ${detail})` : ""),
        isError: true,
      };

    case "gone":
      // Nothing the user or their admin can do, and nothing that retrying will
      // fix. Saying "try again shortly" here would be a lie that costs them
      // time, so this says plainly that it is broken on our side and needs a
      // fix from us.
      return {
        content:
          `Knowva can't fetch ${what} any more: Atlassian has withdrawn the API it uses. This ` +
          "is a fault on our side that needs a code change -- it is not a permissions problem, " +
          "not their sign-in, and not something retrying will fix. Tell them plainly that this " +
          "part is currently broken and has been reported, and do not suggest they try again or " +
          "reconnect. Do not guess at what it would have returned." +
          (detail ? ` (Technical detail, not for the user: ${detail})` : ""),
        isError: true,
      };

    case "rate-limited":
      return {
        content:
          "Confluence is rate-limiting this account. Tell the user to try again in a few minutes.",
        isError: true,
      };

    case "not-found":
      return {
        content:
          `Confluence returned nothing for ${what}. Tell the user it could not be found, and ` +
          "that they may simply not have access to it. Do not assert that it does not exist.",
        isError: true,
      };

    case "transport-error":
    default:
      return {
        content:
          `Knowva couldn't reach Confluence for ${what}. Tell the user it's a problem on our ` +
          "side, not theirs, and to try again shortly. Do not guess at what it would have " +
          "returned." +
          (detail ? ` (Technical detail, not for the user: ${detail})` : ""),
        isError: true,
      };
  }
}
