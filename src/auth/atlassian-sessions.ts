/**
 * What Knowva remembers about a user's Atlassian sign-in.
 *
 * ===========================================================================
 * ONE SIGN-IN, TWO HALVES, WITH DIFFERENT LIFETIMES AND DIFFERENT RISKS.
 *
 * This store used to hold an identity and nothing else, because the Atlassian
 * flow existed only to answer "which Jira account is this Teams user?" and then
 * threw the token away. Confluence changed that: searching Confluence has to
 * happen AS THE USER, because Confluence's own permission model is what makes
 * the results correct. So a token now has to be kept and refreshed.
 *
 * Rather than let that one requirement drag the whole record up to credential
 * status, the two things are stored as separately disposable halves:
 *
 *   identity -- accountId, display name, email, site. NOT a credential. Grants
 *               nothing. Never expires (an accountId is immutable for the life
 *               of an Atlassian account). Used by find_my_xray_items, which
 *               searches on the SHARED Xray key and only needs to know whose
 *               items to look for.
 *
 *   tokens   -- a live OAuth access/refresh pair. A real credential, with a
 *               real expiry, refreshed on demand. Used ONLY by the Confluence
 *               tools, which call Confluence as this user.
 *
 * WHY THAT SPLIT EARNS ITS KEEP, beyond tidiness: the halves fail
 * independently. When a refresh fails -- the user revoked the grant, or 90 days
 * passed -- dropTokens() clears the credential and LEAVES THE IDENTITY. The
 * user is asked to sign in again for Confluence, while find_my_xray_items keeps
 * working, because nothing about "which Jira account is yours" stopped being
 * true. Storing one blob would have logged them out of a tool that never needed
 * the token in the first place.
 *
 * THE AWKWARDNESS THIS DOES NOT SOLVE, and it is worth being honest about:
 * there is one OAuth app and therefore one consent screen. A user who only ever
 * asks about Xray still consents to Confluence read scopes they will never
 * exercise, because the scopes are fixed at authorize time and Knowva cannot
 * know in advance which tools they will use. Splitting that would mean two
 * Atlassian apps and two sign-in flows, which is a worse experience to fix a
 * theoretical over-grant. Flagged rather than hidden -- see ATLASSIAN_SCOPES in
 * src/auth/atlassian.ts.
 * ===========================================================================
 *
 * ---------------------------------------------------------------------------
 * IN-MEMORY ONLY, AND NOW THIS IS A REAL TRADE-OFF RATHER THAN A CHEAP ONE.
 *
 * Same module-level Map as src/llm/history.ts, src/user/email-preferences.ts
 * and src/auth/github-sessions.ts: it dies with the process, it is
 * per-instance, and every deploy or idle recycle clears it.
 *
 * WHAT CHANGED. While this file held only an identity, losing it cost the user
 * one sign-in click and nothing else. Now it holds live OAuth credentials, so
 * every argument in the header of src/auth/github-sessions.ts applies here in
 * full -- in particular:
 *
 *   - A durable store (Redis or otherwise) would hold real Confluence
 *     credentials for every user, and would need encryption at rest, a rotation
 *     story, an access policy, and a delete path wired to sign-out. A backup
 *     that quietly retains tokens after a user revokes them is holding
 *     credentials nobody can see and nobody is expiring.
 *   - Today a revoked grant is dead the moment Atlassian says so, and the
 *     process forgets it on restart. That is the property a durable store gives
 *     up, and it must be replaced deliberately rather than by accident.
 *
 * Until then the cost is visible and bounded: users re-connect after each
 * deploy. Losing a token fails safe.
 * ---------------------------------------------------------------------------
 *
 * KEYED ON THE ENTRA OBJECT ID, not `activity.from.id`, for the reason set out
 * at length in src/auth/github-sessions.ts: a surface-specific id can differ
 * between where the user signed in and where they later ask, and this map must
 * answer "who is this?" identically in both places. With a live credential in
 * the record, a wrong key would mean spending one person's Confluence access on
 * another person's question.
 */

/** Who a Teams user is, over in Atlassian. Not a credential; grants nothing. */
export interface AtlassianIdentity {
  /**
   * The Atlassian account id. Immutable for the life of the account, and the
   * only field find_my_xray_items depends on.
   */
  accountId: string;
  /** Display name, used for the "mentions me" text search and for messages. */
  displayName?: string;
  /**
   * Email address. OFTEN ABSENT, and that is normal rather than a failure:
   * Atlassian hides it unless the user's profile visibility allows it.
   */
  emailAddress?: string;
  /** e.g. https://knowva.atlassian.net -- used to build issue and page links. */
  siteUrl?: string;
  /**
   * The cloud id this identity was resolved against.
   *
   * REUSED BY CONFLUENCE RATHER THAN RE-RESOLVED. It is the same value in
   * https://api.atlassian.com/ex/jira/{cloudId}/... and
   * https://api.atlassian.com/ex/confluence/{cloudId}/... -- one site, one
   * cloud id, both products. Resolving it a second time for Confluence would be
   * a round trip to learn a fact already in hand.
   */
  cloudId?: string;
  /** Epoch ms the identity was resolved. */
  resolvedAt: number;
}

/** A live OAuth credential for one user. This half IS sensitive. */
export interface AtlassianTokens {
  accessToken: string;
  /**
   * ROTATES ON EVERY REFRESH. Atlassian issues a new refresh token each time
   * and invalidates the one just spent, so this must be overwritten from the
   * latest token response -- never carried forward. A store that forgets does
   * not degrade; it breaks permanently on the second refresh.
   */
  refreshToken?: string;
  /** Epoch ms when the access token stops working (Atlassian issues ~1 hour). */
  expiresAt?: number;
  /**
   * Epoch ms of the last successful token operation. Atlassian expires refresh
   * tokens after 90 days of INACTIVITY rather than at a fixed age, and its
   * token response carries no `refresh_token_expires_in`, so this is the clock.
   */
  lastUsedAt: number;
  /**
   * The scopes this token was ACTUALLY granted, from the token response's
   * `scope` field -- which is not necessarily what ATLASSIAN_SCOPES asked for.
   *
   * ===========================================================================
   * THIS EXISTS BECAUSE ITS ABSENCE COST A DEBUGGING SESSION.
   *
   * Atlassian fixes a token's scopes at CONSENT time. Adding scopes to the app
   * in the developer console does not upgrade tokens already issued, and
   * refreshing does not pick them up either -- the refresh token carries the
   * original consent. Atlassian's own docs are explicit: "users who previously
   * consented to the scopes will need to re-consent to the new scopes."
   *
   * The failure mode that produces is nasty precisely because it is partial. A
   * session minted when the app only had `read:jira-user` keeps working for
   * Jira forever, while every Confluence call returns 401 "scope does not
   * match" -- which reads exactly like a revoked grant, and sends whoever is
   * debugging it to check a revocation that never happened.
   *
   * Recording what was granted turns that into one obvious log line at sign-in.
   * Undefined means the session predates this field, not that nothing was
   * granted -- treat it as unknown rather than empty.
   * ===========================================================================
   */
  scopes?: string[];
}

/** One user's Atlassian sign-in: always an identity, sometimes a credential. */
export interface AtlassianSession {
  identity: AtlassianIdentity;
  /** Absent once a refresh has failed, or if the grant never included one. */
  tokens?: AtlassianTokens;
}

/** An in-flight sign-in, remembered between the redirect and the callback. */
interface PendingState {
  /** Entra object id of the user who started the flow. */
  userId: string;
  createdAt: number;
}

/** How long a `state` stays valid. Matches github-sessions.ts. */
const STATE_TTL_MS = 10 * 60 * 1000;

/**
 * Refresh this far before real expiry. A token that dies mid-request costs the
 * user their answer; a few minutes of early refresh costs one token call.
 * Mirrors EXPIRY_SKEW_MS in src/auth/github-sessions.ts.
 */
const EXPIRY_SKEW_MS = 5 * 60 * 1000;

/** Atlassian retires a refresh token after 90 days without use. */
const REFRESH_INACTIVITY_MS = 90 * 24 * 60 * 60 * 1000;

const sessions = new Map<string, AtlassianSession>();
const pendingStates = new Map<string, PendingState>();

/** Object ids are GUIDs whose casing is not guaranteed stable between sources. */
function normalize(userId: string): string {
  return userId.trim().toLowerCase();
}

export function getSession(userId: string): AtlassianSession | undefined {
  return sessions.get(normalize(userId));
}

export function setSession(userId: string, session: AtlassianSession): void {
  sessions.set(normalize(userId), session);
}

/** Forgets everything about this user's Atlassian sign-in. Used by sign-out. */
export function clearSession(userId: string): void {
  sessions.delete(normalize(userId));
}

/**
 * The cheap half. Returns who this user is in Atlassian, with no reference to
 * -- and no requirement for -- a live token.
 *
 * This is the ONLY accessor find_my_xray_items uses. Keeping it separate is
 * what lets the Xray path keep working after a Confluence refresh failure.
 */
export function getIdentity(userId: string): AtlassianIdentity | undefined {
  const session = getSession(userId);
  return session ? session.identity : undefined;
}

/**
 * Drops the credential half, keeping the identity.
 *
 * Called when a refresh fails unrecoverably. The user must re-authorize before
 * Confluence works again, but Knowva still knows which Jira account is theirs,
 * so the Xray tool is unaffected.
 */
export function dropTokens(userId: string): void {
  const session = getSession(userId);
  if (session) session.tokens = undefined;
}

/** True when the access token is expired or close enough that it should be refreshed. */
export function isAccessTokenStale(tokens: AtlassianTokens, now = Date.now()): boolean {
  if (tokens.expiresAt === undefined) return false;
  return tokens.expiresAt - EXPIRY_SKEW_MS <= now;
}

/** True when the refresh token has itself expired, so the credential is unrecoverable. */
export function isRefreshTokenExpired(tokens: AtlassianTokens, now = Date.now()): boolean {
  if (!tokens.refreshToken) return true;
  return now - tokens.lastUsedAt > REFRESH_INACTIVITY_MS;
}

/**
 * Records a pending sign-in. The `state` value itself is generated by the
 * caller (crypto.randomBytes in src/auth/atlassian.ts), so this module is never
 * trusted to pick unguessable values.
 */
export function rememberPendingState(state: string, userId: string): void {
  sweepExpiredStates();
  pendingStates.set(state, { userId, createdAt: Date.now() });
}

/**
 * Redeems a `state` from the OAuth callback, returning the user it belongs to.
 *
 * SINGLE USE, exactly as in github-sessions.ts: the entry is deleted whether or
 * not it turns out to be valid, so a replayed callback cannot establish a
 * second session, and an expired one cannot be retried into working.
 */
export function redeemPendingState(state: string): { userId: string } | null {
  const pending = pendingStates.get(state);
  pendingStates.delete(state);

  if (!pending) return null;
  if (Date.now() - pending.createdAt > STATE_TTL_MS) return null;

  return { userId: pending.userId };
}

/** Drops states nobody came back for. Called on write, not on a timer. */
function sweepExpiredStates(now = Date.now()): void {
  for (const [state, pending] of pendingStates) {
    if (now - pending.createdAt > STATE_TTL_MS) pendingStates.delete(state);
  }
}

/** Test/diagnostic seam. Not used in the request path. */
export function pendingStateCount(): number {
  return pendingStates.size;
}
