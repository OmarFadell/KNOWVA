import type { ILogger } from "@microsoft/teams.common";
import config from "../../config";
import { describeError } from "../errors";

/**
 * Talks to Xray Cloud with ONE SHARED SERVICE CREDENTIAL.
 *
 * ===========================================================================
 * READ THIS BEFORE "FIXING" THE FACT THAT THIS IS NOT PER-USER.
 *
 * Every other credential in Knowva is per-user, and the headers of
 * src/auth/github.ts and src/auth/acting-identity.ts argue at length for why.
 * This file is the exception, and it is not an oversight:
 *
 *   XRAY CLOUD HAS NO PER-USER AUTH MODEL AT ALL. Its API is authenticated with
 *   a Client ID / Client Secret pair generated from an API Key in Xray's Global
 *   Settings. There is no OAuth flow, no 3LO, no delegated token, and no way
 *   for a caller to act as somebody else. An app-level credential is the only
 *   credential Xray offers.
 *
 * ---------------------------------------------------------------------------
 * THE CREDENTIAL IS THE ONLY BOUNDARY. THIS CHANGED, AND IT MATTERS.
 *
 * An earlier version of this milestone took an admin-set allowlist of project
 * keys (XRAY_PROJECTS) and would only ever ask about keys on that list. This
 * file used to carry a comment forbidding any discovery call, precisely so the
 * shared credential could never enumerate.
 *
 * THAT ALLOWLIST HAS BEEN REMOVED, DELIBERATELY, and discovery is now the whole
 * mechanism -- see discoverXrayProjects below. The consequence is worth stating
 * plainly rather than leaving for somebody to find out:
 *
 *   Knowva reports every project this API key can see. There is no second
 *   filter. Any Teams user who can reach Knowva can learn which projects have
 *   Xray tests and roughly how many.
 *
 * So the disclosure decision now lives entirely with whoever issues the API
 * key, and the way to narrow what Knowva reports is to issue a key for an Xray
 * user with narrower Jira permissions. There is no application-level knob, and
 * adding one back would mean reintroducing the allowlist.
 * ---------------------------------------------------------------------------
 *
 * Docs:
 *   https://docs.getxray.app/display/XRAYCLOUD/Authentication+-+REST
 *   https://docs.getxray.app/display/XRAYCLOUD/GraphQL+API
 *   https://us.xray.cloud.getxray.app/doc/graphql/gettests.doc.html
 */

/**
 * Fixed, not configurable. Xray Cloud publishes regional hostnames
 * (us./eu./au.xray.cloud.getxray.app) but the apex host routes correctly for
 * every tenant, and one fixed value is one less thing to get wrong per
 * environment.
 */
const XRAY_BASE_URL = "https://xray.cloud.getxray.app";

/** Network budget per Xray call. A Teams turn has a few seconds before the user gives up. */
const REQUEST_TIMEOUT_MS = 15_000;

/** Xray caps `limit` at 100 on every paged query. */
const PAGE_SIZE = 100;

/**
 * How many pages of tests to scan while discovering projects.
 *
 * THIS IS THE COST CEILING, AND IT BOUNDS THE WRONG THING ON PURPOSE. Discovery
 * walks TESTS to find PROJECTS, so the work scales with how many tests the
 * tenant has, not with how many projects -- a tenant with 50,000 tests in three
 * projects would page 500 times to learn three things. Twenty pages (2,000
 * tests) finds every project in any realistic tenant while keeping one turn
 * comfortably inside Xray's rate limit of 300 calls per 5 minutes.
 *
 * When the cap is hit the result says so rather than presenting a partial scan
 * as complete -- see `scannedAll` on XrayDiscovery.
 */
const MAX_DISCOVERY_PAGES = 20;

/**
 * ---------------------------------------------------------------------------
 * TOKEN LIFETIME, CHECKED AGAINST THE DOCS RATHER THAN ASSUMED.
 *
 * Xray's documentation is explicit on two points and silent on a third, and the
 * silence is the important part:
 *
 *   - The bearer token from POST /api/v2/authenticate "expires after 24 hours".
 *   - The API key itself does not expire.
 *   - THERE IS NO REFRESH ENDPOINT AND NO REFRESH TOKEN. Unlike GitHub and
 *     Atlassian there is nothing to exchange for a fresh token: the only way to
 *     renew is to call /authenticate again with the same client id and secret.
 *
 * That last point is why this module caches a token and renews it an hour early
 * rather than implementing a refresh flow -- there is no flow to implement.
 * Renewing early costs one extra HTTP call a day and avoids a token dying
 * mid-turn, which would cost a user their answer.
 * ---------------------------------------------------------------------------
 */
const TOKEN_LIFETIME_MS = 24 * 60 * 60 * 1000;
const TOKEN_RENEWAL_SKEW_MS = 60 * 60 * 1000;

/** Why an Xray call could not produce a result. Each maps to distinct user-facing advice. */
export type XrayFailure =
  | "not-configured"
  | "auth-failed"
  | "unknown-project"
  /**
   * The JQL matched more than 100 issues, which Xray refuses outright. Kept
   * apart from "query-error" because it is the one failure the CALLER can fix,
   * by narrowing the query -- see searchTestsByJql.
   */
  | "jql-too-broad"
  | "query-error"
  | "transport-error";

/** One project discovered from the tests the service credential can see. */
export interface XrayProject {
  /** Jira's numeric project id, as Xray reports it. Always present. */
  projectId: string;
  /** Jira project key, e.g. "ALPHA". Absent if Xray would not resolve it. */
  key?: string;
  /** Jira project name. Absent if Xray would not resolve it. */
  name?: string;
  /** Exact number of Xray Tests in the project, counted separately per project. */
  testCount?: number;
}

export interface XrayDiscovery {
  ok: boolean;
  /** Set when ok. Projects seen, in the order first encountered. */
  projects?: XrayProject[];
  /** Set when ok: the total number of Xray Tests this credential can see. */
  totalTests?: number;
  /**
   * Set when ok. False means the page cap stopped the scan, so a project whose
   * tests all sat beyond it would have been missed. The tool must not present a
   * partial scan as a complete list.
   */
  scannedAll?: boolean;
  /** Set when not ok. */
  reason?: XrayFailure;
  /** Set when not ok, when the far end said something worth relaying. */
  detail?: string;
}

/** One Xray Test returned by a JQL search. */
export interface XrayTestMatch {
  /** Jira issue key, e.g. "ALPHA-42". */
  key: string;
  summary?: string;
  /** Atlassian accountId of the assignee, when the issue has one. */
  assigneeAccountId?: string;
  assigneeName?: string;
}

export interface XrayTestSearch {
  ok: boolean;
  /** Set when ok. */
  matches?: XrayTestMatch[];
  /** Set when ok: how many Tests matched in total, which may exceed `matches`. */
  total?: number;
  /** Set when not ok. */
  reason?: XrayFailure;
  /** Set when not ok. */
  detail?: string;
}

/** True when this deployment has an Xray service credential configured at all. */
export function isXrayConfigured(): boolean {
  return Boolean(config.xray.clientId && config.xray.clientSecret);
}

let cachedToken: { token: string; expiresAt: number } | undefined;

/**
 * Finds every project the service credential can see, with an exact test count
 * for each.
 *
 * ===========================================================================
 * WHY THIS WALKS TESTS RATHER THAN ASKING FOR A PROJECT LIST.
 *
 * XRAY HAS NO "LIST PROJECTS" QUERY. That is the constraint everything here
 * follows from. The candidates and why they do not work:
 *
 *   - getProjectsSettings(projectIdsOrKeys:, limit:, start:) -- its first
 *     argument is the list you are trying to obtain. Paging arguments make it
 *     look like an enumeration, but there is no documented behaviour for
 *     calling it with no ids, so it cannot be relied on to produce a list you
 *     do not already have.
 *   - getTests(jql: "...") -- Xray documents that a JQL returning more than 100
 *     issues errors out, so JQL cannot carry a broad sweep.
 *
 * What does work: getTests with NO filter at all returns tests across
 * everything the credential can see, and each Test carries `projectId` plus a
 * `jira(fields:)` escape hatch that resolves arbitrary Jira fields -- including
 * `project`, which is where the key and name come from. So projects are
 * discovered as a side effect of walking tests. That is genuinely backwards,
 * and it is the only route the API offers.
 *
 * TWO PHASES, BECAUSE THE SAMPLE CANNOT BE TRUSTED FOR COUNTS. Phase one pages
 * tests to learn WHICH projects exist. Phase two asks each discovered project
 * for its own total via getTests(projectId:, limit: 1). Counting occurrences
 * within phase one instead would report the size of the scan, not the size of
 * the project, and would silently undercount the moment the page cap is hit.
 * ===========================================================================
 */
export async function discoverXrayProjects(log: ILogger): Promise<XrayDiscovery> {
  if (!isXrayConfigured()) return { ok: false, reason: "not-configured" };

  const byProjectId = new Map<string, XrayProject>();
  let totalTests = 0;
  let scannedAll = false;
  let scanned = 0;

  // --- Phase 1: walk tests, collect distinct projects ----------------------
  for (let page = 0; page < MAX_DISCOVERY_PAGES; page++) {
    const result = await graphql<{
      getTests?: { total?: number; results?: Array<Record<string, unknown>> };
    }>(
      "query KnowvaDiscoverProjects($limit: Int!, $start: Int!) {\n" +
        "  getTests(limit: $limit, start: $start) {\n" +
        "    total\n" +
        "    results { projectId jira(fields: [\"project\"]) }\n" +
        "  }\n" +
        "}",
      { limit: PAGE_SIZE, start: page * PAGE_SIZE },
      log
    );

    if (!result.ok) return { ok: false, reason: result.reason, detail: result.detail };

    const getTests = result.data ? result.data.getTests : undefined;
    const results = getTests && Array.isArray(getTests.results) ? getTests.results : [];
    if (typeof getTests?.total === "number") totalTests = getTests.total;

    for (const test of results) {
      const project = extractProject(test);
      if (!project) continue;
      if (!byProjectId.has(project.projectId)) byProjectId.set(project.projectId, project);
    }

    scanned += results.length;

    // Short page, empty page, or we have now seen everything Xray reports.
    if (results.length < PAGE_SIZE || scanned >= totalTests) {
      scannedAll = true;
      break;
    }
  }

  log.info(
    `xray: discovery scanned ${scanned} of ${totalTests} test(s) and found ` +
      `${byProjectId.size} project(s)` +
      (scannedAll ? "" : ` -- stopped at the ${MAX_DISCOVERY_PAGES}-page cap`)
  );

  // --- Phase 2: an exact test count per discovered project -----------------
  const projects = [...byProjectId.values()];
  for (const project of projects) {
    const count = await countTestsInProject(project.projectId, log);
    // A per-project count failing is not worth failing the whole list over --
    // the project was definitely discovered, so it definitely has tests. Leave
    // testCount undefined and let the tool report the project without a number.
    if (count.ok) project.testCount = count.total;
  }

  return { ok: true, projects, totalTests, scannedAll };
}

/**
 * Exact number of Xray Tests in one project.
 *
 * `projectId` is Jira's NUMERIC project id. The schema documents this argument
 * as "the id of the project", not the key, so a key is not valid input here --
 * which is exactly why discovery has to produce ids rather than keys.
 *
 * limit: 1 with only `total` selected is one resolver and one node, comfortably
 * inside Xray's limits of 25 resolvers and 10,000 nodes per call.
 */
async function countTestsInProject(
  projectId: string,
  log: ILogger
): Promise<{ ok: boolean; total?: number }> {
  const result = await graphql<{ getTests?: { total?: number } }>(
    "query KnowvaProjectTestCount($projectId: String!) {\n" +
      "  getTests(projectId: $projectId, limit: 1) { total }\n" +
      "}",
    { projectId },
    log
  );

  if (!result.ok) return { ok: false };

  const total = result.data && result.data.getTests ? result.data.getTests.total : undefined;
  return typeof total === "number" ? { ok: true, total } : { ok: false };
}

/**
 * Pulls a project out of one Test result.
 *
 * DEFENSIVE ABOUT SHAPE ON PURPOSE. The schema renders `projectId` as `[String]`
 * and `jira(fields:)` as `[JSON]!`, while every documented example uses both as
 * plain scalars/objects. Rather than bet on which the server actually sends,
 * both forms are accepted. A missing key or name is survivable -- the id is
 * what the second phase needs -- so this returns a project as long as it can
 * find an id at all.
 */
function extractProject(test: Record<string, unknown>): XrayProject | null {
  const projectId = firstScalar(test.projectId);
  if (!projectId) return null;

  const jira = firstObject(test.jira);
  const project = jira && typeof jira.project === "object" ? (jira.project as any) : undefined;

  return {
    projectId,
    key: project && typeof project.key === "string" ? project.key : undefined,
    name: project && typeof project.name === "string" ? project.name : undefined,
  };
}

function firstScalar(value: unknown): string | undefined {
  const candidate = Array.isArray(value) ? value[0] : value;
  if (typeof candidate === "string" && candidate) return candidate;
  if (typeof candidate === "number") return String(candidate);
  return undefined;
}

function firstObject(value: unknown): Record<string, unknown> | undefined {
  const candidate = Array.isArray(value) ? value[0] : value;
  return candidate && typeof candidate === "object"
    ? (candidate as Record<string, unknown>)
    : undefined;
}

/**
 * Runs a JQL search that returns Xray Tests.
 *
 * ===========================================================================
 * WHY THIS IS THE ONLY SEARCH SHAPE AVAILABLE, AND WHAT IT CANNOT DO.
 *
 * Xray's GraphQL exposes no general Jira issue search. `getTests(jql:)` applies
 * the JQL and then intersects the result with Xray Tests, so it can find Tests
 * matching arbitrary Jira criteria -- assignee, text, project -- but it can
 * only ever hand back Tests.
 *
 * DEFECTS ARE THEREFORE OUT OF REACH FROM THIS FILE, and not because of a
 * missing query shape. Defects are ordinary Jira issues, and the Xray API key
 * authenticates to xray.cloud.getxray.app only -- it is not a Jira credential
 * and cannot call /rest/api/3/search on api.atlassian.com. Finding Defects
 * means holding a Jira token for somebody, which the identity-only design in
 * src/auth/atlassian.ts deliberately does not do. Do not try to route a Defect
 * search through here; there is no route.
 *
 * DO NOT ADD `issuetype = Test` TO THE JQL. getTests already restricts to
 * Tests, and Xray's Test issue type NAME is configurable per project -- the
 * literal string "Test" is a convention, not a guarantee. Adding the clause
 * buys nothing and silently returns zero rows on any project that renamed it.
 * (The same configurability is why defect issue types come from
 * ProjectSettings.defectIssueTypes rather than being assumed.)
 *
 * THE 100-ISSUE CEILING IS REAL AND APPLIES HERE. Xray documents that a JQL
 * returning more than 100 issues is rejected outright, and that cap is on the
 * JQL's own result set rather than on the Tests returned. A per-user query is
 * normally far below it, but "normally" is not "always", so the condition is
 * classified as its own failure (`jql-too-broad`) and passed to the caller to
 * explain, rather than being reported as a generic error the user can do
 * nothing about.
 * ===========================================================================
 */
export async function searchTestsByJql(
  jql: string,
  limit: number,
  log: ILogger
): Promise<XrayTestSearch> {
  if (!isXrayConfigured()) return { ok: false, reason: "not-configured" };

  const result = await graphql<{
    getTests?: { total?: number; results?: Array<Record<string, unknown>> };
  }>(
    `query KnowvaSearchTests($jql: String!, $limit: Int!) {
      getTests(jql: $jql, limit: $limit) {
        total
        results { jira(fields: ["key", "summary", "assignee"]) }
      }
    }`,
    { jql, limit: Math.min(Math.max(limit, 1), PAGE_SIZE) },
    log
  );

  if (!result.ok) return { ok: false, reason: result.reason, detail: result.detail };

  const getTests = result.data ? result.data.getTests : undefined;
  const rows = getTests && Array.isArray(getTests.results) ? getTests.results : [];

  const matches: XrayTestMatch[] = [];
  for (const row of rows) {
    const jira = firstObject(row.jira);
    const key = jira && typeof jira.key === "string" ? jira.key : undefined;
    // Without a key the row cannot be cited or linked, and an uncitable result
    // is worse than no result -- the model would describe an issue the user
    // cannot go and look at.
    if (!key) continue;

    const assignee =
      jira && jira.assignee && typeof jira.assignee === "object"
        ? (jira.assignee as any)
        : undefined;

    matches.push({
      key,
      summary: jira && typeof jira.summary === "string" ? jira.summary : undefined,
      assigneeAccountId:
        assignee && typeof assignee.accountId === "string" ? assignee.accountId : undefined,
      assigneeName:
        assignee && typeof assignee.displayName === "string" ? assignee.displayName : undefined,
    });
  }

  return {
    ok: true,
    matches,
    total: typeof getTests?.total === "number" ? getTests.total : matches.length,
  };
}

/**
 * Confirms one project has Xray data, accepting a project KEY or a numeric id.
 *
 * NOT USED BY THE ACTIVE TOOL. Kept for the parked per-user tool in
 * src/tools/jira-xray-projects.ts, which starts from a project list obtained
 * from Jira and needs to check specific projects rather than discover them.
 * getProjectSettings is the one query documented to accept either form, which
 * is what lets that caller pass whichever it has.
 */
export async function checkXrayProject(
  projectIdOrKey: string,
  log: ILogger
): Promise<{ ok: boolean; projectId?: string; testCount?: number; reason?: XrayFailure; detail?: string }> {
  if (!isXrayConfigured()) return { ok: false, reason: "not-configured" };

  const settings = await graphql<{ getProjectSettings?: { projectId?: string } }>(
    "query KnowvaProjectSettings($projectIdOrKey: String!) {\n" +
      "  getProjectSettings(projectIdOrKey: $projectIdOrKey) { projectId }\n" +
      "}",
    { projectIdOrKey },
    log
  );

  if (!settings.ok) return { ok: false, reason: settings.reason, detail: settings.detail };

  const projectId =
    settings.data && settings.data.getProjectSettings
      ? settings.data.getProjectSettings.projectId
      : undefined;

  if (!projectId) {
    log.warn(`xray: no Xray project settings for ${projectIdOrKey}`);
    return { ok: false, reason: "unknown-project" };
  }

  const count = await countTestsInProject(projectId, log);
  if (!count.ok) {
    return { ok: false, reason: "query-error", detail: "no total in the GraphQL response" };
  }

  return { ok: true, projectId, testCount: count.total };
}

interface GraphqlResult<T> {
  ok: boolean;
  data?: T;
  reason?: XrayFailure;
  detail?: string;
}

/** One authenticated GraphQL round trip, with the error classification callers need. */
async function graphql<T>(
  query: string,
  variables: Record<string, unknown>,
  log: ILogger
): Promise<GraphqlResult<T>> {
  let token: string;
  try {
    token = await getServiceToken(log);
  } catch (err) {
    log.error("xray: could not authenticate the service credential", describeError(err), err);
    return { ok: false, reason: "auth-failed", detail: messageOf(err) };
  }

  let response: Response;
  try {
    response = await fetch(`${XRAY_BASE_URL}/api/v2/graphql`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Accept: "application/json",
      },
      body: JSON.stringify({ query, variables }),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    });
  } catch (err) {
    log.error("xray: GraphQL request failed", describeError(err), err);
    return { ok: false, reason: "transport-error", detail: messageOf(err) };
  }

  // A 401 here means the token we just minted was rejected, which in practice
  // means the API key was revoked or rotated between authenticate and query.
  // Dropping the cache makes the next call re-authenticate rather than keep
  // presenting a token the far end has retired.
  if (response.status === 401 || response.status === 403) {
    cachedToken = undefined;
    log.error(`xray: GraphQL returned HTTP ${response.status} -- service credential rejected`);
    return { ok: false, reason: "auth-failed", detail: `HTTP ${response.status}` };
  }

  if (!response.ok) {
    return {
      ok: false,
      reason: "transport-error",
      detail: `HTTP ${response.status} ${response.statusText}`,
    };
  }

  const body = (await response.json().catch(() => ({}))) as {
    data?: T;
    errors?: Array<{ message?: string }>;
  };

  // GraphQL reports query-level failures with HTTP 200 and an `errors` array,
  // so the status code alone is not the error path -- the same trap
  // src/auth/github.ts documents for GitHub's token endpoint, in a different
  // protocol.
  if (body.errors && body.errors.length > 0) {
    const detail = body.errors
      .map((e) => e.message)
      .filter(Boolean)
      .join("; ");
    log.warn(`xray: GraphQL error -- ${detail || "(no message)"}`);

    // Xray's own words for the 100-issue ceiling are "refine the jql search".
    // Matching on the message is unlovely, but the alternative is reporting a
    // fixable, extremely common condition as a generic query failure -- and
    // this is the one error whose remedy the caller can actually act on.
    if (/refine the jql|more than 100 issues/i.test(detail)) {
      return { ok: false, reason: "jql-too-broad", detail };
    }

    const looksLikeMissingProject = /not found|does not exist|no project/i.test(detail);
    return {
      ok: false,
      reason: looksLikeMissingProject ? "unknown-project" : "query-error",
      detail,
    };
  }

  return { ok: true, data: body.data };
}

/**
 * Mints (or reuses) the shared bearer token.
 *
 * Xray's authenticate endpoint answers with a JSON *string* -- the token,
 * wrapped in double quotes -- rather than a JSON object. Parsing it as JSON
 * yields the token directly; the fallback strips the quotes by hand in case a
 * future version returns it as text/plain.
 */
async function getServiceToken(log: ILogger): Promise<string> {
  const now = Date.now();
  if (cachedToken && cachedToken.expiresAt > now) return cachedToken.token;

  log.info("xray: authenticating the shared service credential");

  const response = await fetch(`${XRAY_BASE_URL}/api/v2/authenticate`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "application/json" },
    body: JSON.stringify({
      client_id: config.xray.clientId,
      client_secret: config.xray.clientSecret,
    }),
    signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
  });

  const raw = await response.text();

  if (!response.ok) {
    // Xray's own message names the cause, and it is app-level configuration, so
    // it is safe to log. 401 here is an invalid or revoked API key.
    throw new Error(
      `Xray authenticate returned HTTP ${response.status}` + (raw ? `: ${raw.slice(0, 200)}` : "")
    );
  }

  let token: string;
  try {
    const parsed = JSON.parse(raw);
    token = typeof parsed === "string" ? parsed : String(parsed);
  } catch {
    token = raw.trim().replace(/^"|"$/g, "");
  }

  if (!token) throw new Error("Xray authenticate returned an empty token");

  cachedToken = { token, expiresAt: now + TOKEN_LIFETIME_MS - TOKEN_RENEWAL_SKEW_MS };
  return token;
}

/** Test/diagnostic seam, and the path used when a credential is rotated. */
export function clearCachedToken(): void {
  cachedToken = undefined;
}

function messageOf(err: unknown): string | undefined {
  const e = err as any;
  return e && typeof e.message === "string" ? e.message : undefined;
}
