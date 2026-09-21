import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";

/**
 * Talks to Confluence Cloud AS THE SIGNED-IN USER.
 *
 * ===========================================================================
 * THIS IS THE ONE ATLASSIAN SURFACE WITH REAL PER-USER ENFORCEMENT.
 *
 * Every call here carries a user's own OAuth access token, and Confluence
 * applies that user's real space and page permissions server-side. A page they
 * cannot read does not come back -- not filtered by Knowva, but never returned
 * in the first place.
 *
 * That is worth stating next to the Xray client, which is the opposite case.
 * src/xray/client.ts uses one shared app-level credential because Xray Cloud
 * offers no per-user auth at all, so its results are "what Knowva can see",
 * scoped afterwards to items naming the asker. Confluence needs no such
 * approximation, and the difference is why the two live in separate modules
 * with separate headers rather than behind one "Atlassian client".
 *
 * THE PRACTICAL CONSEQUENCE, which is also the reason there is no space
 * allowlist: there is nothing for Knowva to restrict. Confluence has already
 * decided what this user may see by the time a response arrives, and adding an
 * application-level filter on top could only ever hide things the user is
 * entitled to, never reveal anything. Scope of search is "all spaces the
 * signed-in user can access", enforced by Confluence.
 * ===========================================================================
 *
 * ---------------------------------------------------------------------------
 * WHY v1 ENDPOINTS AND NOT v2, SINCE v2 IS THE NEWER API.
 *
 * CQL SEARCH DOES NOT EXIST IN v2. Confluence's REST API v2 has 29 API groups
 * -- page, blogpost, space, comment, attachment, label, task and so on -- and
 * not one of them is search. Keyword search is GET /wiki/rest/api/search, a v1
 * endpoint, and there is no v2 equivalent.
 *
 * Atlassian also documents that the two scope families should not be mixed
 * across API versions: v1 pairs with classic scopes (read:confluence-content.all),
 * v2 with granular ones (read:page:confluence). Since search is the primary
 * operation and must be v1, using v2 for the page fetch would mean holding both
 * scope families for one feature. So the whole Confluence surface is v1 and the
 * scope set stays classic -- which is Atlassian's own recommendation anyway.
 *
 * If a v2 search endpoint ever ships, moving BOTH calls together is the right
 * migration. Moving one is not.
 * ---------------------------------------------------------------------------
 *
 * Docs:
 *   https://developer.atlassian.com/cloud/confluence/rest/v1/api-group-search/
 *   https://developer.atlassian.com/cloud/confluence/rest/v1/api-group-content/
 *   https://developer.atlassian.com/cloud/confluence/advanced-searching-using-cql/
 */

/** Network budget per Confluence call. A Teams turn has a few seconds. */
const REQUEST_TIMEOUT_MS = 15_000;

/** Why a Confluence call could not produce a result. Each maps to distinct advice. */
export type ConfluenceFailure =
  | "unauthorized"
  /**
   * The token is valid but was never granted the Confluence scopes -- almost
   * always a session created before those scopes were added to the app. Kept
   * apart from `unauthorized` because the remedy differs and the wrong message
   * sends people hunting for a revocation that never happened.
   */
  | "scope-mismatch"
  /**
   * The endpoint itself has been withdrawn (HTTP 410). Distinct from every
   * other failure here because nothing about the request, the token or the
   * user is wrong -- the API is simply gone, and no retry, reconnect or
   * permission change will bring it back. Only a code change will.
   *
   * Atlassian is progressively removing the Confluence REST API v1 Content
   * endpoints in favour of v2, so this is the shape that removal takes.
   */
  | "gone"
  | "forbidden"
  | "not-found"
  | "bad-query"
  | "rate-limited"
  | "transport-error";

/** One search hit. Deliberately a summary -- the body comes from getPage. */
export interface ConfluenceSearchHit {
  /** Content id. This is what get_confluence_page takes. */
  id: string;
  title: string;
  /** "page", "blogpost", "comment", "attachment". */
  type?: string;
  spaceKey?: string;
  spaceName?: string;
  /** ISO 8601 UTC, when the content was last changed. */
  lastModified?: string;
  lastModifiedBy?: string;
  /**
   * Confluence's own highlighted excerpt. Carries <b>...</b> markers around the
   * matched terms, which are stripped before the model sees them.
   */
  excerpt?: string;
  /** Absolute URL to the page. */
  url?: string;
}

export interface ConfluenceSearchResult {
  ok: boolean;
  hits?: ConfluenceSearchHit[];
  /** Total matches Confluence reports, which may exceed the page returned. */
  total?: number;
  reason?: ConfluenceFailure;
  detail?: string;
}

/** One page, with its body rendered to plain text. */
export interface ConfluencePage {
  id: string;
  title: string;
  spaceKey?: string;
  spaceName?: string;
  lastModified?: string;
  lastModifiedBy?: string;
  version?: number;
  /** Body, converted from storage-format XHTML to readable text. */
  text: string;
  /** True when the body was cut to fit a turn. */
  truncated: boolean;
  url?: string;
}

export interface ConfluencePageResult {
  ok: boolean;
  page?: ConfluencePage;
  reason?: ConfluenceFailure;
  detail?: string;
}

/** Longest page body handed to the model, in characters. */
const MAX_BODY_CHARS = 12_000;

/**
 * Searches Confluence with CQL, as the signed-in user.
 *
 * ===========================================================================
 * THE CQL IS BUILT, NOT PASSED THROUGH, AND THAT IS A DELIBERATE BOUNDARY.
 *
 * The tool takes a plain keyword string from the model and this function wraps
 * it in `text ~ "..."`, escaping it on the way. The model never supplies raw
 * CQL.
 *
 * Two reasons, and the second is the one that matters. First, a model writing
 * query-language syntax gets it subtly wrong and the user sees "no results"
 * rather than "bad query". Second and more important: CQL is a query language
 * over a permissioned store, and letting a model compose it freely invites
 * clauses nobody intended. The permission model would still hold -- Confluence
 * enforces per-user access regardless of what the CQL says -- but the blast
 * radius of a confused model should not extend to query construction when a
 * keyword string is all the feature needs.
 *
 * `type = page` is included so search returns pages rather than comments and
 * attachments, which cannot be usefully fetched by get_confluence_page.
 * ===========================================================================
 */
export async function searchConfluence(
  accessToken: string,
  cloudId: string,
  siteUrl: string | undefined,
  query: string,
  limit: number,
  log: ILogger
): Promise<ConfluenceSearchResult> {
  const cql = `type = page AND text ~ ${quoteCql(query)}`;

  const url =
    `${base(cloudId)}/wiki/rest/api/search` +
    `?cql=${encodeURIComponent(cql)}` +
    `&limit=${encodeURIComponent(String(limit))}` +
    // `space` gives the space key and name for attribution; `version` gives the
    // last-modified timestamp and author. Both are what the two-stage pattern
    // needs for a citable snippet, and neither pulls a body.
    `&expand=${encodeURIComponent("content.space,content.version")}`;

  const response = await request(url, accessToken, log);
  if (!response.ok) return { ok: false, reason: response.reason, detail: response.detail };

  const body = response.body as {
    results?: Array<Record<string, any>>;
    totalSize?: number;
  };

  const hits: ConfluenceSearchHit[] = [];
  for (const row of body.results || []) {
    const content = row.content || {};
    if (!content.id) continue;

    hits.push({
      id: String(content.id),
      title: content.title || row.title || "(untitled)",
      type: content.type,
      spaceKey: content.space ? content.space.key : undefined,
      spaceName: content.space ? content.space.name : undefined,
      lastModified: content.version ? content.version.when : row.lastModified,
      lastModifiedBy:
        content.version && content.version.by ? content.version.by.displayName : undefined,
      excerpt: stripHighlight(row.excerpt),
      url: pageUrl(siteUrl, String(content.id)),
    });
  }

  return { ok: true, hits, total: typeof body.totalSize === "number" ? body.totalSize : hits.length };
}

/**
 * Space key and name, keyed by the numeric space id v2 returns.
 *
 * Space metadata is stable and non-sensitive -- a key and a name that everyone
 * who can see the space can see anyway -- so caching it process-wide is safe in
 * a way caching page content would not be. It turns the extra lookup v2 forces
 * (see below) into roughly one call per space for the life of the process
 * rather than one per page read.
 */
const spaceCache = new Map<string, { key?: string; name?: string }>();

/**
 * Fetches one page's full body, as the signed-in user.
 *
 * Stage two of the two-stage pattern the SharePoint and email tools already
 * use: search returns citable summaries, and the body is pulled only for the
 * one page that turns out to matter. That keeps a broad search affordable and
 * means a page nobody needed is never read at all -- which is a privacy
 * property, not only a cost one.
 *
 * ===========================================================================
 * v2, BECAUSE v1 IS GONE -- NOT BECAUSE v2 IS NICER.
 *
 * This used to call GET /wiki/rest/api/content/{id}?expand=body.storage,space,version.
 * Atlassian has now withdrawn it, and says so explicitly:
 *
 *     410 GoneException: This deprecated endpoint has been removed.
 *
 * So there is no choice, and one consequence is worth understanding before
 * anybody "tidies" the extra call below away:
 *
 * v1 RETURNED THE SPACE AND THE EDITOR INLINE. `expand=space,version` gave back
 * `space: {key, name}` and `version.by.displayName` in the same response. v2
 * returns `spaceId` -- a NUMERIC ID, not a key -- and `version.authorId`, an
 * account id. Neither is usable in a citation: "page 2850817 in space 98304,
 * last edited by 712020:abc..." tells a reader nothing.
 *
 * Hence the second call to resolve the space, and hence the editor's display
 * name being DROPPED rather than shown. Resolving the author would mean a third
 * round trip and a third scope (read:user:confluence) to recover one line of
 * attribution, which is not worth it -- so the page result carries no
 * lastModifiedBy and the tool tells the model not to attribute the edit.
 *
 * THIS ALSO FORCES MIXING SCOPE FAMILIES. Search is v1 (classic
 * `search:confluence`); this is v2 (granular `read:page:confluence`,
 * `read:space:confluence`). Atlassian advises against mixing them, and that
 * advice stops being followable the moment they remove the v1 endpoint you
 * would need in order to stay pure. See ATLASSIAN_SCOPES in
 * src/auth/atlassian.ts.
 * ===========================================================================
 */
export async function getConfluencePage(
  accessToken: string,
  cloudId: string,
  siteUrl: string | undefined,
  pageId: string,
  log: ILogger
): Promise<ConfluencePageResult> {
  const url =
    `${base(cloudId)}/wiki/api/v2/pages/${encodeURIComponent(pageId)}` +
    `?body-format=storage&include-version=true`;

  const response = await request(url, accessToken, log);
  if (!response.ok) return { ok: false, reason: response.reason, detail: response.detail };

  const data = response.body as Record<string, any>;
  if (!data || !data.id) return { ok: false, reason: "not-found" };

  const storage = data.body && data.body.storage ? data.body.storage.value : "";
  const text = storageToText(typeof storage === "string" ? storage : "");
  const truncated = text.length > MAX_BODY_CHARS;

  // Best-effort: a page that cannot be attributed to a space is still worth
  // returning, so a failed space lookup degrades the citation rather than
  // failing the read.
  const space = data.spaceId
    ? await resolveSpace(accessToken, cloudId, String(data.spaceId), log)
    : undefined;

  return {
    ok: true,
    page: {
      id: String(data.id),
      title: data.title || "(untitled)",
      spaceKey: space ? space.key : undefined,
      spaceName: space ? space.name : undefined,
      lastModified: data.version ? data.version.createdAt : undefined,
      // Deliberately absent: v2 gives only an account id here. See the header.
      lastModifiedBy: undefined,
      version: data.version ? data.version.number : undefined,
      text: truncated ? `${text.slice(0, MAX_BODY_CHARS)}\n...(truncated)` : text,
      truncated,
      url: webUiUrl(siteUrl, data) || pageUrl(siteUrl, String(data.id)),
    },
  };
}

function base(cloudId: string): string {
  return `https://api.atlassian.com/ex/confluence/${encodeURIComponent(cloudId)}`;
}

/**
 * Turns v2's numeric spaceId into a key and name, via a cached lookup.
 *
 * Returns undefined rather than throwing when the lookup fails. A page whose
 * space cannot be named is still worth reading -- the citation loses the space
 * and keeps the title and link -- so this must not be able to fail the read.
 *
 * NOTE THE NEGATIVE CACHING. A failed lookup is cached as an empty record, so a
 * space this token cannot read (or a transient failure on a space that is
 * repeatedly requested) costs one call rather than one per page read. The cost
 * is that a genuine transient failure sticks until the process restarts, which
 * is the same lifetime as every other cache in this codebase.
 */
async function resolveSpace(
  accessToken: string,
  cloudId: string,
  spaceId: string,
  log: ILogger
): Promise<{ key?: string; name?: string } | undefined> {
  const cached = spaceCache.get(spaceId);
  if (cached) return cached.key || cached.name ? cached : undefined;

  const response = await request(
    `${base(cloudId)}/wiki/api/v2/spaces/${encodeURIComponent(spaceId)}`,
    accessToken,
    log
  );

  if (!response.ok) {
    log.warn(
      `confluence: could not resolve space ${spaceId} (${response.reason}); the page will be ` +
        "cited without its space"
    );
    spaceCache.set(spaceId, {});
    return undefined;
  }

  const data = response.body as Record<string, any>;
  const record = {
    key: typeof data?.key === "string" ? data.key : undefined,
    name: typeof data?.name === "string" ? data.name : undefined,
  };
  spaceCache.set(spaceId, record);
  return record.key || record.name ? record : undefined;
}

/**
 * Builds a page link from v2's own `_links`, which is more reliable than
 * constructing one.
 *
 * `_links.webui` is a path such as /spaces/DOCS/pages/2850817/Some+Title, and
 * `_links.base` is the site's wiki root when present. Preferring what the API
 * hands back avoids guessing at a URL shape that Atlassian may change -- the
 * constructed viewpage.action form stays as the fallback.
 */
function webUiUrl(siteUrl: string | undefined, data: Record<string, any>): string | undefined {
  const links = data._links;
  const webui = links && typeof links.webui === "string" ? links.webui : undefined;
  if (!webui) return undefined;

  const root =
    links && typeof links.base === "string" && links.base
      ? links.base.replace(/\/+$/, "")
      : siteUrl
        ? `${siteUrl.replace(/\/+$/, "")}/wiki`
        : undefined;

  if (!root) return undefined;
  return `${root}${webui.startsWith("/") ? "" : "/"}${webui}`;
}

/**
 * Builds a human-usable page link.
 *
 * Uses the /pages/viewpage.action?pageId= form rather than the prettier
 * /spaces/KEY/pages/ID/Title one, because it needs only the id -- the pretty
 * form needs a URL-slugged title, and a link that is subtly wrong is worse than
 * one that is plain. Confluence redirects this to the canonical URL.
 */
function pageUrl(siteUrl: string | undefined, pageId: string): string | undefined {
  if (!siteUrl) return undefined;
  return `${siteUrl.replace(/\/+$/, "")}/wiki/pages/viewpage.action?pageId=${encodeURIComponent(pageId)}`;
}

/**
 * Escapes a value for a CQL string literal.
 *
 * Backslashes first, or escaping the quotes would then be undone by the
 * backslash pass. Same ordering as the JQL quoting in
 * src/tools/find-my-xray-items.ts, and for the same reason.
 */
function quoteCql(value: string): string {
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}"`;
}

interface RawResponse {
  ok: boolean;
  body?: unknown;
  reason?: ConfluenceFailure;
  detail?: string;
}

/**
 * One authenticated GET, with the error classification both callers need.
 *
 * NOTE THE 401 HANDLING, OR RATHER THE LACK OF RETRY. A 401 here means
 * Confluence rejected a token the auth layer believed was live -- almost always
 * a revoked grant. It is reported as `unauthorized` and the tool turns that
 * into a fresh sign-in offer. Retrying or refreshing here would duplicate logic
 * that belongs in src/auth/atlassian.ts, and would hide a revocation behind a
 * retry loop.
 */
async function request(url: string, accessToken: string, log: ILogger): Promise<RawResponse> {
  let response: Response;
  try {
    response = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" },
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
  } catch (err) {
    log.error("confluence: request failed", describeError(err), err);
    return { ok: false, reason: "transport-error", detail: messageOf(err) };
  }

  if (response.ok) {
    try {
      return { ok: true, body: await response.json() };
    } catch (err) {
      return { ok: false, reason: "transport-error", detail: "response was not JSON" };
    }
  }

  const detail = (await response.text().catch(() => "")).slice(0, 300);

  switch (response.status) {
    case 401:
      // ===================================================================
      // A 401 HERE HAS TWO QUITE DIFFERENT CAUSES AND THEY MUST NOT BE
      // CONFLATED. This log line used to assert revocation, which cost real
      // debugging time when the cause was a scope shortfall.
      //
      //   "scope does not match" -- the token is perfectly valid and the grant
      //       is intact; it simply was not granted the Confluence scopes.
      //       Happens to every session created before the Confluence scopes
      //       were added to the app, because Atlassian fixes scopes at consent
      //       time and neither a refresh nor a console change upgrades an
      //       existing token. The giveaway is that Jira keeps working.
      //
      //   anything else -- the user revoked the grant, or it genuinely expired.
      //
      // Atlassian's own body text distinguishes them, which is why `detail` is
      // now logged instead of being dropped on the floor. Never remove it: it
      // is the only thing in the whole pipeline that says WHY.
      // ===================================================================
      if (/scope\s*does\s*not\s*match|scope/i.test(detail)) {
        log.warn(
          "confluence: 401 SCOPE MISMATCH -- this token was not granted the Confluence scopes. " +
            "The grant is intact; it predates those scopes being added. The user must " +
            `re-consent (/jira-signout then reconnect). Atlassian said: ${detail}`
        );
        return { ok: false, reason: "scope-mismatch", detail };
      }

      log.warn(`confluence: 401 -- Atlassian rejected the token. Atlassian said: ${detail}`);
      return { ok: false, reason: "unauthorized", detail };
    case 403:
      return { ok: false, reason: "forbidden", detail };
    case 404:
      return { ok: false, reason: "not-found", detail };

    case 410:
      // ===================================================================
      // THE ENDPOINT IS GONE. Logged at error rather than warn, and named
      // explicitly, because this previously fell through to the default case
      // and was reported as "Knowva couldn't reach Confluence" -- which is
      // actively misleading. Confluence was reached perfectly well; it
      // answered, and its answer was that this API no longer exists.
      //
      // The URL is logged because WHICH endpoint died is the entire content
      // of this error, and it is the difference between a one-line fix and a
      // day of looking in the wrong place.
      // ===================================================================
      log.error(
        `confluence: 410 GONE -- Atlassian has withdrawn this endpoint: ${url.split("?")[0]}. ` +
          "This needs a code change to a supported API; no retry or reconnect will help. " +
          `Atlassian said: ${detail}`
      );
      return { ok: false, reason: "gone", detail };
    case 400:
      // Almost always a malformed CQL string. Since the CQL is built here
      // rather than supplied by the model, this points at a bug in the builder
      // or at a genuinely unsearchable input.
      log.warn(`confluence: 400 on ${url.split("?")[0]} -- ${detail}`);
      return { ok: false, reason: "bad-query", detail };
    case 429:
      return { ok: false, reason: "rate-limited", detail };
    default:
      log.error(`confluence: HTTP ${response.status} ${response.statusText} -- ${detail}`);
      return {
        ok: false,
        reason: "transport-error",
        detail: `HTTP ${response.status} ${response.statusText}`,
      };
  }
}

/**
 * Converts Confluence storage format to readable plain text.
 *
 * Storage format is XHTML with Confluence-specific macro elements, and the goal
 * here is a faithful-enough rendering for a model to read and quote from -- not
 * a perfect one.
 *
 * THE ORDER OF OPERATIONS MATTERS. Macro and structural elements that carry no
 * readable text (ac:parameter, layout wrappers) are dropped whole, INCLUDING
 * their contents, before the generic tag strip. Doing it the other way round
 * would leave macro parameter values -- layout hints, panel colours, user keys
 * -- inlined into the prose as though somebody had written them.
 *
 * Block boundaries become newlines before tags are stripped, or every paragraph
 * and table cell would run together into one unreadable line.
 */
function storageToText(storage: string): string {
  if (!storage) return "";

  let text = storage;

  // Whole elements whose contents are markup, not prose.
  text = text.replace(/<ac:parameter[\s\S]*?<\/ac:parameter>/g, "");
  text = text.replace(/<ri:[^>]*\/>/g, "");
  text = text.replace(/<!--[\s\S]*?-->/g, "");

  // Structural boundaries -> newlines, before the tags disappear.
  text = text.replace(/<\/(p|div|li|tr|h[1-6]|blockquote)>/gi, "\n");
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/t[dh]>/gi, "\t");

  // Everything else.
  text = text.replace(/<[^>]+>/g, "");

  // Entities. Ampersand LAST, or "&amp;lt;" would decode to "<".
  text = text
    .replace(/&nbsp;/g, " ")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");

  // Collapse the blank lines all that tag-stripping leaves behind.
  return text
    .split("\n")
    .map((line) => line.trim())
    .filter((line, index, all) => line.length > 0 || (index > 0 && all[index - 1].length > 0))
    .join("\n")
    .trim();
}

/** Confluence wraps matched terms in <b>...</b> (and @@@hl@@@ markers in some versions). */
function stripHighlight(excerpt: unknown): string | undefined {
  if (typeof excerpt !== "string" || !excerpt) return undefined;
  return excerpt
    .replace(/@@@hl@@@|@@@endhl@@@/g, "")
    .replace(/<[^>]+>/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function messageOf(err: unknown): string | undefined {
  const e = err as any;
  return e && typeof e.message === "string" ? e.message : undefined;
}
