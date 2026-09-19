/**
 * What Knowva remembers about a user's GitHub authorization.
 *
 * Two stores, both keyed on the asking user, both deliberately small:
 *
 *   1. SESSIONS -- the access/refresh token pair a completed OAuth flow
 *      produced, keyed by Entra object id.
 *   2. PENDING STATES -- the short-lived `state` values handed to GitHub when a
 *      sign-in starts, so the callback can prove the code it was given belongs
 *      to a flow we actually initiated, and work out who for.
 *
 * ===========================================================================
 * WHY THE KEY IS THE ENTRA OBJECT ID, NOT `activity.from.id`
 *
 * The milestone brief says "keyed by Teams user ID", and there are two
 * candidates for that. `activity.from.id` is a Teams-surface identifier that
 * varies by channel and carries no meaning outside the Bot Framework.
 * `activity.from.aadObjectId` is the user's Entra object id -- the same value
 * `/me` returns as `id`, and the same key space the email opt-out in
 * src/user/email-preferences.ts already uses.
 *
 * The object id wins, and the reason is not tidiness. A GitHub session is a
 * live credential for somebody's real GitHub account. Looking one up under a
 * key that can differ between the surface where the user signed in and the
 * surface where they later ask a question would, at best, silently lose the
 * session -- and at worst hand it to the wrong person. Using the identity that
 * src/auth/acting-identity.ts resolves from the Graph token itself means the
 * lookup key is derived from an authenticated fact, not from a routing detail.
 * ===========================================================================
 *
 * ---------------------------------------------------------------------------
 * IN-MEMORY ONLY, AND THIS ONE IS A GENUINE TRADE-OFF -- NOT AN OVERSIGHT.
 *
 * Same shape as src/llm/history.ts and src/user/email-preferences.ts: a
 * module-level Map that dies with the process. Every deploy, every idle recycle
 * on App Service, and it is per-instance, so a second worker sees none of it.
 *
 * What that costs here is specific: every user has to click through GitHub
 * sign-in again after each restart. Not a security problem -- losing a token
 * fails safe, unlike the email opt-out, which fails *open* when it is forgotten
 * -- but it is the most user-visible of the three, because it is the only one
 * where the user has to do something to recover.
 *
 * REDIS IS THE OBVIOUS REPLACEMENT AND IT IS NOT A FREE WIN. Weigh it properly
 * before adopting it:
 *
 *   For:     sessions survive deploys and scale-out, which is the difference
 *            between "sign in once" and "sign in after every release". Also the
 *            only way the 6-month refresh token is worth anything -- an 8-hour
 *            access token in a store that resets on deploy never gets old
 *            enough to need refreshing.
 *   Against: it puts live OAuth credentials for every user in a network service
 *            that now needs encryption at rest, a rotation story, an access
 *            policy, and a backup policy that must NOT quietly retain tokens
 *            after a user revokes them. Today a revoked GitHub authorization is
 *            dead the moment GitHub says so and the process forgets it on
 *            restart; a durable store has to be told, and if it is not, it is
 *            holding credentials nobody can see and nobody is expiring.
 *
 * That second column is why this is flagged rather than decided here. If
 * sessions do go durable, they should go durable with an explicit TTL at the
 * storage layer and a delete path wired to sign-out -- not as a side effect of
 * "we added Redis for conversation history".
 * ---------------------------------------------------------------------------
 */

/** A completed GitHub authorization for one user. */
export interface GitHubSession {
  accessToken: string;
  /**
   * Present because the GitHub App has "expire user tokens" enabled. Without
   * it a session simply dies at `expiresAt` and the user signs in again.
   */
  refreshToken?: string;
  /** Epoch ms when the access token stops working. Undefined = non-expiring. */
  expiresAt?: number;
  /** Epoch ms when the refresh token itself expires (~6 months out). */
  refreshTokenExpiresAt?: number;
  /** GitHub login, purely so messages can say *which* account is connected. */
  login?: string;
  /** Epoch ms the session was first established. */
  createdAt: number;
}

/** An in-flight sign-in, remembered between the redirect and the callback. */
interface PendingState {
  /** Entra object id of the user who started the flow. */
  userId: string;
  createdAt: number;
}

/**
 * How long a `state` stays valid. Long enough to log into GitHub and approve,
 * short enough that a leaked authorize URL is not a standing invitation.
 */
const STATE_TTL_MS = 10 * 60 * 1000;

/**
 * Refresh this far before real expiry. A token that dies mid-request costs the
 * user their answer; a few minutes of early refresh costs one token call.
 * Mirrors EXPIRY_SKEW_MS in src/graph/app-client.ts.
 */
const EXPIRY_SKEW_MS = 5 * 60 * 1000;

const sessions = new Map<string, GitHubSession>();
const pendingStates = new Map<string, PendingState>();

/** Object ids are GUIDs whose casing is not guaranteed stable between sources. */
function normalize(userId: string): string {
  return userId.trim().toLowerCase();
}

export function getSession(userId: string): GitHubSession | undefined {
  return sessions.get(normalize(userId));
}

export function setSession(userId: string, session: GitHubSession): void {
  sessions.set(normalize(userId), session);
}

export function clearSession(userId: string): void {
  sessions.delete(normalize(userId));
}

/** True when the access token is expired or close enough that it should be refreshed. */
export function isAccessTokenStale(session: GitHubSession, now = Date.now()): boolean {
  if (session.expiresAt === undefined) return false;
  return session.expiresAt - EXPIRY_SKEW_MS <= now;
}

/** True when the refresh token has itself expired, so the session is unrecoverable. */
export function isRefreshTokenExpired(session: GitHubSession, now = Date.now()): boolean {
  if (!session.refreshToken) return true;
  if (session.refreshTokenExpiresAt === undefined) return false;
  return session.refreshTokenExpiresAt <= now;
}

/**
 * Records a pending sign-in and returns the opaque `state` to send to GitHub.
 *
 * The state value itself is the caller's to generate -- it comes from
 * crypto.randomBytes in src/auth/github.ts, so this module never has to be
 * trusted to pick unguessable values.
 */
export function rememberPendingState(state: string, userId: string): void {
  sweepExpiredStates();
  pendingStates.set(state, { userId, createdAt: Date.now() });
}

/**
 * Redeems a `state` from the OAuth callback, returning the user it belongs to.
 *
 * SINGLE USE. The entry is deleted whether or not it turns out to be valid, so
 * a replayed callback cannot establish a second session, and an expired one
 * cannot be retried into working.
 */
export function redeemPendingState(state: string): { userId: string } | null {
  const pending = pendingStates.get(state);
  pendingStates.delete(state);

  if (!pending) return null;
  if (Date.now() - pending.createdAt > STATE_TTL_MS) return null;

  return { userId: pending.userId };
}

/**
 * Drops states nobody came back for. Called on write rather than on a timer:
 * this map only grows when somebody starts a sign-in, so there is no need for
 * an interval that keeps the process awake.
 */
function sweepExpiredStates(now = Date.now()): void {
  for (const [state, pending] of pendingStates) {
    if (now - pending.createdAt > STATE_TTL_MS) pendingStates.delete(state);
  }
}

/** Test/diagnostic seam. Not used in the request path. */
export function pendingStateCount(): number {
  return pendingStates.size;
}
