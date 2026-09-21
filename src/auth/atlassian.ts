import { randomBytes } from "crypto";
import type { ILogger } from "@microsoft/teams.common";
import config from "../../config";
import { describeError } from "../errors";
import { escapeHtml, oauthResultPage } from "./oauth-result-page";
import {
  clearSession,
  dropTokens,
  getIdentity,
  getSession,
  isAccessTokenStale,
  isRefreshTokenExpired,
  redeemPendingState,
  rememberPendingState,
  setSession,
  type AtlassianIdentity,
  type AtlassianTokens,
} from "./atlassian-sessions";

/**
 * The Atlassian OAuth 2.0 (3LO) flow, serving TWO consumers with two very
 * different relationships to the token it produces.
 *
 * ===========================================================================
 * ONE SIGN-IN, TWO USES. KEEPING THESE STRAIGHT IS THE POINT OF THIS FILE.
 *
 *   1. IDENTITY, for find_my_xray_items.
 *      Resolves the user's Atlassian accountId once, via /rest/api/3/myself.
 *      The Xray search itself runs on the SHARED Xray service credential, not
 *      on this token, because Xray Cloud has no per-user auth model at all.
 *      That tool reads getAtlassianIdentity() and never touches a token.
 *
 *   2. ACCESS, for search_confluence and get_confluence_page.
 *      Confluence is different in the way that matters: its REST API enforces
 *      the signed-in user's real page and space permissions. So calls there are
 *      made AS THE USER, with a retained, refreshed access token, and the
 *      results are genuinely per-user rather than an approximation of one.
 *
 * WHY THAT DISTINCTION IS WORTH THE COMPLEXITY. It would be simpler to hold a
 * token and use it for everything, or to hold none and approximate. Neither is
 * honest here: Xray cannot enforce per-user access however much we would like
 * it to, and Confluence can and does. Collapsing the two would mean either
 * claiming a guarantee Xray cannot make, or throwing away one Confluence does.
 *
 * THE TOKEN AND THE IDENTITY ARE STORED AS SEPARATELY DISPOSABLE HALVES (see
 * src/auth/atlassian-sessions.ts). A failed refresh drops the credential and
 * keeps the identity, so a user whose Confluence session dies can still use the
 * Xray tool without signing in again. Do not merge them back into one blob.
 * ===========================================================================
 *
 * ---------------------------------------------------------------------------
 * THIS FILE USED TO DISCARD THE TOKEN, AND THE HISTORY EXPLAINS THE SHAPE.
 *
 * While Xray identity was the only consumer, completeSignIn() called /myself
 * and let the token fall out of scope -- no storage, no refresh, and no
 * `offline_access` scope, because there was nothing to refresh. Confluence
 * required that to change.
 *
 * What came back with it: rotating refresh tokens, a 90-day inactivity clock,
 * and getValidAccessToken(). What did NOT change: the Xray path, which still
 * only ever reads the identity half. If you are adding a third consumer, decide
 * which half it belongs to before writing any code.
 * ---------------------------------------------------------------------------
 *
 * STRUCTURALLY THIS MIRRORS src/auth/github.ts, now more closely than before:
 * buildSignInUrl -> getValidAccessToken -> completeSignIn -> signOut ->
 * exchange -> registerAtlassianOAuthRoutes. If you change one of them, check
 * whether the other needs the same change. A plain Adaptive Card with a link
 * rather than a Bot Framework OAuth card, a browser GET callback on this app's
 * own HTTP server rather than a `signin/tokenExchange` invoke, and a single-use
 * server-side `state` -- that file's header explains why none of this can route
 * through the `graph` OAuth connection.
 *
 * WHERE ATLASSIAN STILL DIFFERS FROM GITHUB, all load-bearing:
 *
 *   1. THE CLOUD ID STEP, which GitHub has no equivalent of. An Atlassian token
 *      addresses nothing by itself: every call goes to
 *      https://api.atlassian.com/ex/<product>/<cloudId>/... and the cloud id is
 *      only discoverable by asking. It is resolved once at sign-in and cached
 *      on the identity -- ONE cloud id serves both products, so Confluence
 *      reuses the value Jira resolved rather than asking again.
 *   2. ROTATING REFRESH TOKENS. Atlassian invalidates the refresh token it just
 *      accepted and issues a new one. The new value MUST overwrite the old or
 *      the session breaks permanently on the second refresh.
 *   3. ONE CALLBACK URL, NOT A LIST. A GitHub App accepts several; an Atlassian
 *      OAuth 2.0 integration has a single field, so a recreated dev tunnel
 *      means REPLACING it. That is why PERSISTENT_TUNNEL_HOST matters here.
 *   4. JSON, NOT FORM ENCODING, AND HONEST STATUS CODES. GitHub answers a
 *      failed exchange with HTTP 200 plus an `error` field; Atlassian takes
 *      JSON and returns a real 4xx.
 *
 * Docs:
 *   https://developer.atlassian.com/cloud/jira/platform/oauth-2-3lo-apps/
 *   https://developer.atlassian.com/cloud/confluence/scopes-for-oauth-2-3LO-and-forge-apps/
 */

const ATLASSIAN_AUTHORIZE_URL = "https://auth.atlassian.com/authorize";
const ATLASSIAN_TOKEN_URL = "https://auth.atlassian.com/oauth/token";
const ACCESSIBLE_RESOURCES_URL = "https://api.atlassian.com/oauth/token/accessible-resources";

/**
 * The scopes requested at authorize time.
 *
 * ===========================================================================
 * ALL-GRANULAR, SPANNING TWO API GENERATIONS.
 *
 * Atlassian publishes two scope families -- classic (search:confluence) and
 * granular (read:content-details:confluence) -- and advises against mixing
 * them. This set is entirely granular, which keeps that advice satisfied even
 * though the endpoints behind it span both API versions:
 *
 *   - SEARCH is v1: GET /wiki/rest/api/search. CQL search exists nowhere else
 *     -- REST v2 has 29 API groups and not one of them is search. Its granular
 *     scope is read:content-details:confluence (the classic equivalent, which
 *     this app no longer requests, is search:confluence).
 *   - PAGE FETCH is v2: GET /wiki/api/v2/pages/{id}. It had to move there,
 *     because Atlassian withdrew the v1 Content API, which now answers:
 *
 *         410 GoneException: This deprecated endpoint has been removed.
 *
 * So one product, two API generations, one scope family. Any future "cleanup"
 * that assumes a single API version will break whichever half it forgets.
 *
 *   read:jira-user                  -- GET /rest/api/3/myself, to resolve which
 *                                      Atlassian account belongs to this Teams
 *                                      user. Used by find_my_xray_items only.
 *   read:content-details:confluence -- GET /wiki/rest/api/search (CQL).
 *   read:content:confluence         -- content bodies and summaries.
 *   read:content.property:confluence,
 *   read:content.restriction:confluence,
 *   read:content.metadata:confluence
 *                                   -- content properties, restrictions and
 *                                      metadata the search leg expands.
 *   read:page:confluence            -- GET /wiki/api/v2/pages/{id}. This is
 *                                      what replaced the withdrawn v1 fetch.
 *   read:space:confluence           -- GET /wiki/api/v2/spaces/{id}. v2 returns
 *                                      a page's space as a NUMERIC ID, so
 *                                      naming the space in a citation takes a
 *                                      second lookup.
 *   offline_access                  -- issue a refresh_token. REQUIRED since
 *                                      the session became long-lived.
 *
 * NOT INCLUDED, deliberately: read:user:confluence. v2 reports a page's last
 * editor as an account id, and turning that into a display name would cost a
 * third round trip and a third scope to recover one line of attribution. The
 * page result therefore carries no editor name and the tool tells the model not
 * to attribute the edit. Add this scope if that attribution is ever wanted
 * back.
 *
 * ONE CONSENT SCREEN FOR TWO VERY DIFFERENT USES. A user who only ever asks
 * about Xray still consents to the Confluence scopes, because scopes are fixed
 * at authorize time and Knowva cannot know which tools they will go on to use.
 * Splitting it would mean two Atlassian apps and two sign-ins to avoid a
 * theoretical over-grant, which is a worse trade. See the header of
 * src/auth/atlassian-sessions.ts.
 *
 * These must ALSO be added to the app in the Atlassian developer console
 * (Permissions -> add the Confluence API, then its scopes); asking for a scope
 * the app does not have configured fails the authorize call outright.
 * ===========================================================================
 */
export const ATLASSIAN_SCOPES = [
  "read:jira-user",
  "read:content:confluence",
  "read:content-details:confluence",
  "read:content.property:confluence",
  "read:content.restriction:confluence",
  "read:content.metadata:confluence",
  // v2 page-fetch leg. Added when Atlassian withdrew the v1 Content API and
  // get_confluence_page had to move to /wiki/api/v2/pages/{id}.
  "read:page:confluence",
  "read:space:confluence"
] as const;

/** Path the OAuth callback lands on. Must match the app's single Callback URL. */
export const ATLASSIAN_CALLBACK_PATH = "/auth/atlassian/callback";

/** Network timeout for Atlassian's token and identity endpoints. */
const TOKEN_REQUEST_TIMEOUT_MS = 10_000;

/** True when this deployment has an Atlassian OAuth app configured at all. */
export function isAtlassianConfigured(): boolean {
  return Boolean(
    config.atlassian.clientId && config.atlassian.clientSecret && config.atlassian.oauthOrigin
  );
}

/**
 * The resolved Atlassian identity for a Teams user, or undefined if they have
 * not signed in.
 *
 * THE CHEAP HALF, and the only thing find_my_xray_items needs. Deliberately
 * makes no reference to a token and cannot fail because one expired: an
 * accountId does not stop being true when a credential does.
 */
export function getAtlassianIdentity(userId: string): AtlassianIdentity | undefined {
  return getIdentity(userId);
}

/**
 * Forgets a user's Atlassian sign-in entirely -- identity and token both.
 *
 * DOES NOT REVOKE THE GRANT AT ATLASSIAN. Only the user can do that, from their
 * Atlassian account settings, and the sign-out message says so rather than
 * implying Knowva has done more than forget its own copy. Same honesty as the
 * GitHub sign-out in app.ts.
 */
export function signOut(userId: string): boolean {
  const had = Boolean(getSession(userId));
  clearSession(userId);
  return had;
}

/** Why a Confluence-capable token could not be produced. Mirrors GitHubTokenFailure. */
export type AtlassianTokenFailure =
  | "not-configured"
  | "no-session"
  | "refresh-failed"
  | "session-expired";

/** One shape with optional fields -- see the note on ActingIdentity for why. */
export interface AtlassianTokenResult {
  ok: boolean;
  /** Set when ok: the user's live Atlassian access token. */
  token?: string;
  /** Set when ok: the cloud id to address, reused from sign-in. */
  cloudId?: string;
  /** Set when ok, if known: the site's base URL, for building page links. */
  siteUrl?: string;
  /** Set when not ok. */
  reason?: AtlassianTokenFailure;
}

/**
 * Returns a usable access token for this user, refreshing first if needed.
 *
 * THE CREDENTIAL HALF. Only the Confluence tools call this.
 *
 * Same policy as the GitHub version: a failed refresh is treated as an invalid
 * session rather than an error worth throwing, because the token is gone and
 * the user needs to authorize again -- a normal thing to happen. The difference
 * is WHAT GETS CLEARED. dropTokens() removes the credential and KEEPS the
 * identity, so a Confluence session dying does not silently log the user out of
 * find_my_xray_items, which never needed the token.
 *
 * Atlassian access tokens last about an hour rather than GitHub's eight, so
 * this path runs often -- which is exactly why the rotating refresh token must
 * be written back correctly on every pass.
 */
export async function getValidAccessToken(
  userId: string,
  log: ILogger
): Promise<AtlassianTokenResult> {
  if (!isAtlassianConfigured()) return { ok: false, reason: "not-configured" };

  const session = getSession(userId);
  if (!session || !session.tokens) return { ok: false, reason: "no-session" };

  const identity = session.identity;
  const tokens = session.tokens;

  if (!isAccessTokenStale(tokens)) {
    return {
      ok: true,
      token: tokens.accessToken,
      cloudId: identity.cloudId,
      siteUrl: identity.siteUrl,
    };
  }

  if (isRefreshTokenExpired(tokens)) {
    // Ninety days without using Knowva's Confluence tools, or a grant the user
    // has since revoked from their Atlassian account settings.
    log.info(`atlassian-auth: refresh token for ${userId} has expired; credential dropped`);
    dropTokens(userId);
    return { ok: false, reason: "session-expired" };
  }

  log.info(`atlassian-auth: access token for ${userId} is stale; refreshing`);

  try {
    const refreshed = await exchange(
      { grant_type: "refresh_token", refresh_token: tokens.refreshToken as string },
      log
    );

    // The whole response replaces the stored credential. NOT a merge: the
    // refresh_token that came back is new and the one just spent is dead.
    session.tokens = toTokens(refreshed);
    setSession(userId, session);

    // A refresh does NOT pick up scopes added to the app since the user
    // consented -- the refresh token carries the original grant. So this warns
    // on every refresh of an under-scoped session, rather than once at sign-in
    // where it could scroll away unnoticed.
    warnOnScopeShortfall(userId, refreshed.scope, "refresh", log);

    return {
      ok: true,
      token: session.tokens.accessToken,
      cloudId: identity.cloudId,
      siteUrl: identity.siteUrl,
    };
  } catch (err) {
    log.error(`atlassian-auth: refresh failed for ${userId}`, describeError(err), err);
    dropTokens(userId);
    return { ok: false, reason: "refresh-failed" };
  }
}

/** Turns a token response into the stored credential half. */
function toTokens(tokens: TokenResponse): AtlassianTokens {
  const now = Date.now();
  return {
    accessToken: tokens.access_token,
    refreshToken: tokens.refresh_token || undefined,
    expiresAt: tokens.expires_in ? now + tokens.expires_in * 1000 : undefined,
    lastUsedAt: now,
    // Atlassian returns the granted scopes space-separated. Recorded rather
    // than assumed equal to ATLASSIAN_SCOPES -- see the field's comment in
    // src/auth/atlassian-sessions.ts for the failure this catches.
    scopes: typeof tokens.scope === "string" && tokens.scope ? tokens.scope.split(/\s+/) : undefined,
  };
}

/**
 * Scopes this app now asks for that the user's stored token was NOT granted.
 *
 * ===========================================================================
 * THE CHECK THAT TURNS A CONFUSING 401 INTO A CLEAR ONE.
 *
 * Atlassian fixes a token's scopes at consent time. Adding a scope in the
 * developer console does not upgrade existing tokens, and a refresh does not
 * either -- the refresh token carries the original consent. So a session minted
 * before the Confluence scopes were added keeps working for Jira and returns
 * 401 on every Confluence call, indefinitely, with no self-healing path.
 *
 * Comparing what was granted against what is now asked for detects that
 * exactly, and lets the Confluence tools say "your connection predates
 * Confluence access, please reconnect" instead of "your grant was revoked" --
 * which is what it looks like from the HTTP status alone, and is wrong.
 *
 * Returns [] when the session records no scopes, which means it predates this
 * field rather than that it was granted nothing. Unknown must not be reported
 * as a mismatch, or every pre-existing session would be told to reconnect on a
 * guess.
 * ===========================================================================
 */
/**
 * Logs the scopes a token was granted, loudly when they fall short.
 *
 * The happy path is one info line, which is worth having on its own: it is the
 * only record anywhere of what a session can actually do. The shortfall path is
 * a warning that names the missing scopes and the remedy, because the symptom
 * it predicts -- a 401 from one product while another keeps working -- reads
 * like a revoked grant and will otherwise be debugged as one.
 */
function warnOnScopeShortfall(
  userId: string,
  granted: string | undefined,
  when: string,
  log: ILogger
): void {
  if (!granted) {
    log.warn(
      `atlassian-auth: the token response for ${userId} carried no scope field (${when}), so ` +
        "Knowva cannot tell what this session was granted."
    );
    return;
  }

  const grantedScopes = new Set(granted.split(/\s+/).filter(Boolean));
  const missing = ATLASSIAN_SCOPES.filter((scope) => !grantedScopes.has(scope));

  if (missing.length === 0) {
    log.info(`atlassian-auth: ${userId} granted all requested scopes (${when}) -- ${granted}`);
    return;
  }

  log.warn(
    `atlassian-auth: SCOPE SHORTFALL for ${userId} (${when}). Granted: ${granted}. ` +
      `MISSING: ${missing.join(", ")}. Atlassian fixes scopes at consent time, and neither a ` +
      "refresh nor adding scopes in the developer console upgrades an existing token -- the " +
      "user must re-consent. Until they do, calls needing the missing scopes will return 401 " +
      '"scope does not match" while other products keep working normally.'
  );
}

/**
 * A human-readable summary of one user's stored Atlassian session, for the
 * /jira-debug command.
 *
 * ===========================================================================
 * WHY THIS EXISTS: ONE 401 HAS TWO CAUSES THAT LOOK IDENTICAL FROM OUTSIDE.
 *
 * Confluence answers both of these with the same 401 "scope does not match":
 *
 *   A. The token genuinely lacks the Confluence scopes -- a session that
 *      predates them being added, which no refresh will ever fix.
 *   B. The token HAS every scope we asked for, and Confluence refuses anyway --
 *      which is not a consent problem at all, and means the endpoint itself is
 *      unavailable to this app.
 *
 * The remedies are opposites: (A) reconnect, (B) stop trying to reconnect and
 * change the endpoint. Guessing wrong costs a debugging session, so this prints
 * the one fact that separates them: what the token was ACTUALLY granted.
 *
 * NEVER PRINTS THE TOKEN, only its shape and expiry. A debug command that leaks
 * a live credential into a chat transcript would be a worse bug than the one it
 * is diagnosing.
 * ===========================================================================
 */
export function describeAtlassianSession(userId: string): string {
  const session = getSession(userId);
  if (!session) {
    return "No Atlassian session -- you haven't connected, or the bot restarted since you did.";
  }

  const { identity, tokens } = session;
  const lines = [
    `Atlassian account: ${identity.accountId}` +
      (identity.displayName ? ` (${identity.displayName})` : ""),
    `Site: ${identity.siteUrl || "unknown"}`,
    `Cloud id: ${identity.cloudId ? "present" : "MISSING -- reconnect needed"}`,
  ];

  if (!tokens) {
    lines.push(
      "Token: none. The identity is cached (so Xray still works) but Confluence needs a reconnect."
    );
    return lines.join("\n");
  }

  lines.push(
    `Token expires: ${tokens.expiresAt ? new Date(tokens.expiresAt).toISOString() : "unknown"}`,
    `Refresh token: ${tokens.refreshToken ? "present" : "ABSENT -- offline_access may be missing"}`
  );

  if (!tokens.scopes) {
    lines.push(
      "Granted scopes: NOT RECORDED. This session predates scope recording, so reconnect to " +
        "get a definitive answer."
    );
    return lines.join("\n");
  }

  const missing = ATLASSIAN_SCOPES.filter((scope) => !tokens.scopes?.includes(scope));

  lines.push(`Granted scopes: ${tokens.scopes.join(", ")}`);
  lines.push(
    missing.length > 0
      ? `MISSING: ${missing.join(", ")} -- reconnect to re-consent.`
      : "All requested scopes granted. If Confluence still returns 401, the problem is NOT " +
        "consent -- the endpoint itself is refusing this app."
  );

  return lines.join("\n");
}

export function missingAtlassianScopes(userId: string): string[] {
  const session = getSession(userId);
  if (!session || !session.tokens || !session.tokens.scopes) return [];

  const granted = new Set(session.tokens.scopes);
  return ATLASSIAN_SCOPES.filter((scope) => !granted.has(scope));
}

/**
 * Drops the credential half, keeping the identity.
 *
 * Exposed for the Confluence tools: when Confluence rejects a token the auth
 * layer believed was live, the token cannot be made to work by retrying, so it
 * is discarded and the next question offers a fresh sign-in. The identity
 * survives, so find_my_xray_items is unaffected -- see the header of
 * src/auth/atlassian-sessions.ts.
 */
export function dropAtlassianTokens(userId: string): void {
  dropTokens(userId);
}

export function callbackUrl(): string {
  return `${config.atlassian.oauthOrigin}${ATLASSIAN_CALLBACK_PATH}`;
}

/**
 * Builds the Atlassian authorize URL for one user and remembers the `state`.
 *
 * `state` is 32 bytes of CSPRNG output, single-use, ten-minute TTL, and carries
 * no user data -- it is a lookup key into a server-side map, so an intercepted
 * authorize URL reveals nothing about who it was for. Same as GitHub.
 *
 * Two parameters here have no GitHub counterpart and are both required by
 * Atlassian: `audience=api.atlassian.com` names the API family the token is
 * for, and `prompt=consent` forces the grant screen to be shown. Omitting
 * either produces an authorize URL that fails in a way that reads like a
 * misconfigured client id.
 */
export function buildSignInUrl(userId: string): string | null {
  if (!isAtlassianConfigured()) return null;

  const state = randomBytes(32).toString("base64url");
  rememberPendingState(state, userId);

  const url = new URL(ATLASSIAN_AUTHORIZE_URL);
  url.searchParams.set("audience", "api.atlassian.com");
  url.searchParams.set("client_id", config.atlassian.clientId);
  url.searchParams.set("scope", ATLASSIAN_SCOPES.join(" "));
  url.searchParams.set("redirect_uri", callbackUrl());
  url.searchParams.set("state", state);
  url.searchParams.set("response_type", "code");
  url.searchParams.set("prompt", "consent");
  return url.toString();
}

/**
 * Completes the OAuth flow: code -> token -> cloud id -> /myself -> session.
 *
 * Returns the user the session was established for, so the route can log it,
 * or null when the state was unknown, replayed, or expired.
 *
 * BOTH HALVES ARE ESTABLISHED HERE, and the identity half is established FIRST
 * ON PURPOSE. Resolving the cloud id and account id needs the token anyway, and
 * doing it now means find_my_xray_items never has to make a Jira call of its
 * own -- it reads a cached fact. The credential half is then stored alongside
 * it for the Confluence tools.
 *
 * An earlier version of this function deliberately let the token fall out of
 * scope at this point. It no longer can: Confluence enforces per-user
 * permissions and therefore has to be called as the user.
 */
export async function completeSignIn(
  code: string,
  state: string,
  log: ILogger
): Promise<{ userId: string; identity: AtlassianIdentity } | null> {
  const pending = redeemPendingState(state);
  if (!pending) {
    // Not necessarily an attack -- a bookmarked callback URL, a double-clicked
    // link, or someone slower than ten minutes all land here.
    log.warn("atlassian-auth: OAuth callback presented an unknown, replayed, or expired state");
    return null;
  }

  const tokens = await exchange(
    {
      grant_type: "authorization_code",
      code,
      redirect_uri: callbackUrl(),
    },
    log
  );

  // THE STEP GITHUB HAS NO EQUIVALENT OF. An Atlassian token addresses nothing
  // by itself: every Jira call goes to /ex/jira/{cloudId}/..., and the cloud id
  // is only discoverable by asking. Calling the site's own hostname with this
  // token fails in a way that reads like a permissions problem rather than a
  // routing one.
  const site = await getAccessibleJiraSite(tokens.access_token, log);
  if (!site) {
    throw new Error(
      "This Atlassian account granted access to no Jira site, so there is no account to resolve."
    );
  }

  const me = await fetchMyself(tokens.access_token, site.id, log);

  const identity: AtlassianIdentity = {
    accountId: me.accountId,
    displayName: me.displayName,
    emailAddress: me.emailAddress,
    siteUrl: site.url,
    // Cached here so Confluence never re-resolves it: one cloud id addresses
    // both /ex/jira/ and /ex/confluence/ for the same site.
    cloudId: site.id,
    resolvedAt: Date.now(),
  };

  setSession(pending.userId, { identity, tokens: toTokens(tokens) });

  log.info(
    `atlassian-auth: session established for ${pending.userId} -- accountId ${me.accountId}` +
      (me.displayName ? ` (${me.displayName})` : "") +
      (site.url ? ` on ${site.url}` : "") +
      (me.emailAddress ? "" : " [no email address; Atlassian profile visibility hides it]") +
      (tokens.expires_in
        ? `, access token expires ${new Date(Date.now() + tokens.expires_in * 1000).toISOString()}`
        : "") +
      (tokens.refresh_token ? "" : " [NO refresh token -- offline_access may be missing]")
  );

  // Log what was ACTUALLY granted, not what was asked for. A mismatch here is
  // the difference between "Confluence works" and "Confluence 401s forever",
  // and it is invisible until something fails at the far end.
  warnOnScopeShortfall(pending.userId, tokens.scope, "sign-in", log);

  return { userId: pending.userId, identity };
}

/** One Jira or Confluence site the user has granted access to. */
interface AccessibleResource {
  /** The cloud id. This is what goes in /ex/jira/{cloudId}/... */
  id: string;
  /** e.g. https://knowva.atlassian.net */
  url?: string;
  name?: string;
  scopes?: string[];
}

/**
 * Resolves which Jira site this token can address.
 *
 * WHEN THERE IS MORE THAN ONE SITE this picks the first Jira-bearing one and
 * logs the fact. For identity resolution that matters much less than it did
 * when this flow fetched data: an Atlassian accountId is the same value across
 * every site on the account, so picking the "wrong" one still yields the right
 * account id. It only affects `siteUrl`, and therefore which host the /browse/
 * links point at -- which is worth knowing about if somebody with two sites
 * reports links that land in the wrong Jira.
 */
async function getAccessibleJiraSite(
  accessToken: string,
  log: ILogger
): Promise<AccessibleResource | null> {
  const response = await fetch(ACCESSIBLE_RESOURCES_URL, {
    headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" },
    signal: AbortSignal.timeout(TOKEN_REQUEST_TIMEOUT_MS),
  });

  if (!response.ok) {
    throw new Error(
      `Atlassian accessible-resources returned HTTP ${response.status} ${response.statusText}`
    );
  }

  const resources = (await response.json()) as AccessibleResource[];
  if (!Array.isArray(resources) || resources.length === 0) return null;

  // Confluence-only sites appear in exactly the same list and would 404 every
  // Jira call made against them.
  const jiraSites = resources.filter((r) =>
    (r.scopes || []).some((scope) => scope.indexOf("jira") !== -1)
  );

  const chosen = jiraSites.length > 0 ? jiraSites[0] : resources[0];
  if (!chosen || !chosen.id) return null;

  if (jiraSites.length > 1) {
    log.warn(
      `atlassian-auth: this account has ${jiraSites.length} accessible Jira sites; using ` +
        `${chosen.url || chosen.id} for issue links. The accountId is identical across sites, ` +
        "so only link hostnames are affected."
    );
  }

  return chosen;
}

interface MyselfResponse {
  accountId: string;
  displayName?: string;
  emailAddress?: string;
}

/**
 * GET /rest/api/3/myself -- the one and only reason this OAuth flow exists.
 *
 * ON THE MISSING EMAIL ADDRESS. Atlassian omits `emailAddress` unless the
 * user's profile visibility permits it, and for many tenants it is hidden by
 * default. That is NORMAL, not a failure: the caller must work without it, and
 * the search simply drops its email clause. Treating a missing email as an
 * error here would turn a routine privacy setting into a broken sign-in.
 */
async function fetchMyself(
  accessToken: string,
  cloudId: string,
  log: ILogger
): Promise<MyselfResponse> {
  const response = await fetch(
    `https://api.atlassian.com/ex/jira/${encodeURIComponent(cloudId)}/rest/api/3/myself`,
    {
      headers: { Authorization: `Bearer ${accessToken}`, Accept: "application/json" },
      signal: AbortSignal.timeout(TOKEN_REQUEST_TIMEOUT_MS),
    }
  );

  if (!response.ok) {
    const body = await response.text().catch(() => "");
    throw new Error(
      `Jira /myself returned HTTP ${response.status} ${response.statusText}` +
        (body ? `: ${body.slice(0, 200)}` : "")
    );
  }

  const data = (await response.json()) as MyselfResponse;

  if (!data.accountId) {
    // Without this there is no identity, and every downstream match would be a
    // guess. Fail the sign-in rather than storing a half-identity.
    throw new Error("Jira /myself returned no accountId");
  }

  if (!data.emailAddress) {
    log.info(
      "atlassian-auth: /myself returned no emailAddress (profile visibility). The " +
        "mentions-me search will match on display name only."
    );
  }

  return data;
}

interface TokenResponse {
  access_token: string;
  /** Present when `offline_access` was granted. ROTATES on every refresh. */
  refresh_token?: string;
  expires_in?: number;
  scope?: string;
  token_type?: string;
  error?: string;
  error_description?: string;
}

/**
 * POSTs to Atlassian's token endpoint.
 *
 * JSON BODY, NOT FORM ENCODING, AND HONEST STATUS CODES -- the two places
 * Atlassian differs from GitHub on the wire. GitHub's token endpoint takes
 * application/x-www-form-urlencoded and answers a FAILED exchange with HTTP 200
 * plus an `error` field, which is why src/auth/github.ts checks the body rather
 * than the status. Atlassian takes JSON and returns a real 4xx. The body is
 * still inspected below, but as belt-and-braces rather than the primary error
 * path.
 */
async function exchange(grant: Record<string, string>, log: ILogger): Promise<TokenResponse> {
  const response = await fetch(ATLASSIAN_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "application/json" },
    body: JSON.stringify({
      client_id: config.atlassian.clientId,
      client_secret: config.atlassian.clientSecret,
      ...grant,
    }),
    signal: AbortSignal.timeout(TOKEN_REQUEST_TIMEOUT_MS),
  });

  const data = (await response.json().catch(() => ({}))) as TokenResponse;

  if (!response.ok || data.error || !data.access_token) {
    // error_description names the actual cause (invalid_grant,
    // unauthorized_client, ...), which is the difference between "the user was
    // slow" and "the app is misconfigured".
    log.error(
      `atlassian-auth: token exchange rejected (HTTP ${response.status}) -- ` +
        `${data.error || "no access_token in response"}` +
        (data.error_description ? `: ${data.error_description}` : "")
    );
    throw new Error(data.error || `Atlassian token endpoint returned HTTP ${response.status}`);
  }

  return data;
}

/**
 * The minimum of Express's request/response shapes that this route needs.
 * Declared locally for the same reason as in src/auth/github.ts: express is a
 * transitive dependency here and ships no types of its own.
 */
interface CallbackRequest {
  query: Record<string, unknown>;
}

interface CallbackResponse {
  status(code: number): CallbackResponse;
  type(contentType: string): CallbackResponse;
  send(body: string): unknown;
}

interface RouteRegistrar {
  get(path: string, handler: (req: any, res: any) => void): unknown;
}

/**
 * Mounts the OAuth callback on the app's own HTTP server, alongside
 * /api/messages and the GitHub callback.
 *
 * DELIBERATELY OUTSIDE THE TEAMS AUTH MIDDLEWARE, for the same reason and with
 * the same justification as the GitHub route: the caller is a browser following
 * a redirect from auth.atlassian.com and has no Teams token to present. The
 * route trusts nothing except a `state` this process generated, stored
 * server-side, and will accept only once within ten minutes. Who the identity
 * belongs to comes from that server-side entry, never from the query string.
 */
export function registerAtlassianOAuthRoutes(router: RouteRegistrar, log: ILogger): void {
  if (!isAtlassianConfigured()) {
    log.warn(
      "atlassian-auth: no Atlassian OAuth app configured (ATLASSIAN_CLIENT_ID / " +
        "ATLASSIAN_CLIENT_SECRET / a resolvable public origin), so the OAuth callback route " +
        "was not mounted and find_my_xray_items will report itself unavailable."
    );
    return;
  }

  router.get(ATLASSIAN_CALLBACK_PATH, async (req: CallbackRequest, res: CallbackResponse) => {
    const code = typeof req.query.code === "string" ? req.query.code : "";
    const state = typeof req.query.state === "string" ? req.query.state : "";
    const oauthError = typeof req.query.error === "string" ? req.query.error : "";

    // The user pressed Cancel on Atlassian's grant screen. Not our failure, and
    // it should not read like one.
    if (oauthError) {
      log.info(`atlassian-auth: user did not complete authorization (${oauthError})`);
      res
        .status(200)
        .type("text/html")
        .send(
          oauthResultPage(
            "Jira was not connected",
            "You cancelled the sign-in, so nothing changed. You can close this tab and try " +
              "again from Teams whenever you like."
          )
        );
      return;
    }

    if (!code || !state) {
      res
        .status(400)
        .type("text/html")
        .send(
          oauthResultPage(
            "Something went wrong",
            "That sign-in link was incomplete. Please start again from Teams."
          )
        );
      return;
    }

    try {
      const result = await completeSignIn(code, state, log);

      if (!result) {
        res
          .status(400)
          .type("text/html")
          .send(
            oauthResultPage(
              "That link has expired",
              "Sign-in links are only good for a few minutes and can be used once. Please ask " +
                "Knowva again in Teams to get a fresh one."
            )
          );
        return;
      }

      const who = result.identity.displayName
        ? ` as ${escapeHtml(result.identity.displayName)}`
        : "";
      res
        .status(200)
        .type("text/html")
        .send(
          oauthResultPage(
            "Jira connected",
            `Knowva now knows which Jira account is yours${who}. It doesn't keep any access to ` +
              "your Jira — just your account details, so it can spot items that name you. You " +
              "can close this tab and go back to Teams."
          )
        );
    } catch (err) {
      log.error("atlassian-auth: OAuth callback failed", describeError(err), err);
      res
        .status(500)
        .type("text/html")
        .send(
          oauthResultPage(
            "Something went wrong",
            "Knowva couldn't finish connecting to Jira. That's a problem on our side -- please " +
              "try again from Teams in a moment."
          )
        );
    }
  });

  log.info(`atlassian-auth: identity-only OAuth callback mounted at ${callbackUrl()}`);
}
