import { randomBytes } from "crypto";
import type { ILogger } from "@microsoft/teams.common";
import config from "../../config";
import { describeError } from "../errors";
import {
  clearSession,
  getSession,
  isAccessTokenStale,
  isRefreshTokenExpired,
  redeemPendingState,
  rememberPendingState,
  setSession,
  type GitHubSession,
} from "./github-sessions";

/**
 * The GitHub OAuth flow: per-user authorization for the search_github tool.
 *
 * ===========================================================================
 * THIS IS A SECOND, SEPARATE IDENTITY SYSTEM. DO NOT CONFLATE IT WITH TEAMS SSO.
 *
 * Knowva already has an OAuth flow: `ctx.signin()`, the Bot Framework OAuth
 * connection named `graph`, silently exchanging a Teams token for a Microsoft
 * Graph one. None of that applies here. That machinery is bound to an Azure Bot
 * Service OAuth connection configured against Entra; GitHub is not wired into
 * it, cannot be, and should not be.
 *
 * The practical consequences, each of which is a thing somebody will otherwise
 * try to "fix":
 *
 *   - Sign-in is a PLAIN ADAPTIVE CARD WITH A LINK, not a Bot Framework OAuth
 *     card. An OAuth card would make Teams try to drive an exchange against the
 *     `graph` connection, which knows nothing about GitHub.
 *   - The callback arrives as an ordinary browser GET on this app's own HTTP
 *     server, not as a `signin/tokenExchange` invoke activity.
 *   - Nothing about this is silent. The user clicks a link, GitHub asks them to
 *     authorize, and they come back. That is the point: GitHub is enforcing
 *     *their* repository permissions, which is the whole reason this is
 *     per-user rather than one shared PAT.
 *   - Signing out of Teams does not sign the user out of GitHub, and vice
 *     versa. Two systems, two sessions.
 *
 * src/llm/prompt.ts says the same thing in the model's own terms, so Knowva
 * does not tell a user to "sign in" and leave them guessing which sign-in.
 * ===========================================================================
 *
 * WHY PER-USER OAUTH RATHER THAN A SHARED TOKEN. A single installation-wide PAT
 * would let any Teams user query anything that token can reach, regardless of
 * their own GitHub access -- a textbook confused deputy, where Knowva's
 * permissions get borrowed by whoever asks. With per-user tokens GitHub applies
 * the asking user's real repository access to every query, and Knowva never
 * holds authority that any individual user does not already have.
 *
 * ---------------------------------------------------------------------------
 * CALLBACK URL AND THE DEV TUNNEL. The redirect_uri sent to GitHub must exactly
 * match a Callback URL registered on the GitHub App. Locally that URL contains
 * the dev-tunnel hostname, which changes every time the tunnel is recreated --
 * so a fresh tunnel means adding the new callback URL to KnowvaGithubApp before
 * sign-in will work. GitHub Apps accept several callback URLs, so the practical
 * habit is to add each new tunnel rather than replace the old one.
 * ---------------------------------------------------------------------------
 *
 * Docs:
 *   https://docs.github.com/en/apps/creating-github-apps/authenticating-with-a-github-app/generating-a-user-access-token-for-a-github-app
 *   https://docs.github.com/en/apps/creating-github-apps/authenticating-with-a-github-app/refreshing-user-access-tokens
 */

const GITHUB_AUTHORIZE_URL = "https://github.com/login/oauth/authorize";
const GITHUB_TOKEN_URL = "https://github.com/login/oauth/access_token";
const GITHUB_USER_URL = "https://api.github.com/user";

/** Path the OAuth callback lands on. Must match the GitHub App's Callback URL. */
export const GITHUB_CALLBACK_PATH = "/auth/github/callback";

/** Network timeout for GitHub's token endpoint. A hung sign-in is worse than a failed one. */
const TOKEN_REQUEST_TIMEOUT_MS = 10_000;

/** Why a token could not be produced. Each maps to a distinct user-facing message. */
export type GitHubTokenFailure =
  | "not-configured"
  | "no-session"
  | "refresh-failed"
  | "session-expired";

/** One shape with optional fields -- see the note on ActingIdentity for why. */
export interface GitHubTokenResult {
  ok: boolean;
  /** Set when ok: the user's live GitHub access token. */
  token?: string;
  /** Set when ok, if known: the GitHub login the token belongs to. */
  login?: string;
  /** Set when not ok. */
  reason?: GitHubTokenFailure;
}

/**
 * True when this deployment has a GitHub App configured at all.
 *
 * Modelled on the appClientSecret check in src/graph/app-client.ts: a missing
 * GitHub App is a deployment choice, not a crash. The tool reports it in plain
 * language and everything else keeps working.
 */
export function isGitHubConfigured(): boolean {
  return Boolean(
    config.github.clientId && config.github.clientSecret && config.github.oauthOrigin
  );
}

/** Human-readable account for messages, e.g. "signed in as octocat". */
export function connectedLogin(userId: string): string | undefined {
  return getSession(userId)?.login;
}

/**
 * Builds the GitHub authorize URL for one user and remembers the `state`.
 *
 * `state` is 32 bytes of CSPRNG output, single-use, and expires in ten minutes
 * (see src/auth/github-sessions.ts). It carries no user data itself -- it is a
 * lookup key into a server-side map, so an intercepted authorize URL reveals
 * nothing about who it was for.
 */
export function buildSignInUrl(userId: string): string | null {
  if (!isGitHubConfigured()) return null;

  const state = randomBytes(32).toString("base64url");
  rememberPendingState(state, userId);

  const url = new URL(GITHUB_AUTHORIZE_URL);
  url.searchParams.set("client_id", config.github.clientId);
  url.searchParams.set("redirect_uri", callbackUrl());
  url.searchParams.set("state", state);
  // No `scope` parameter: a GitHub App's user token is bounded by the App's own
  // installed permissions (Contents, Issues, Pull requests -- all read-only),
  // not by OAuth scopes. Asking for scopes here would be meaningless at best.
  return url.toString();
}

export function callbackUrl(): string {
  return `${config.github.oauthOrigin}${GITHUB_CALLBACK_PATH}`;
}

/**
 * Returns a usable access token for this user, refreshing first if needed.
 *
 * Refresh failure is treated as an invalid session rather than an error worth
 * throwing: the token is gone, the user needs to sign in again, and that is a
 * normal thing to have happen after eight hours. The session is cleared so the
 * next call reports "no-session" and offers a sign-in card, instead of
 * repeatedly trying a refresh token GitHub has already rejected.
 */
export async function getValidAccessToken(
  userId: string,
  log: ILogger
): Promise<GitHubTokenResult> {
  if (!isGitHubConfigured()) return { ok: false, reason: "not-configured" };

  const session = getSession(userId);
  if (!session) return { ok: false, reason: "no-session" };

  if (!isAccessTokenStale(session)) {
    return { ok: true, token: session.accessToken, login: session.login };
  }

  if (isRefreshTokenExpired(session)) {
    // Six months without using Knowva, or a token GitHub has since revoked.
    // Distinct from a refresh that failed for some other reason, because the
    // user-facing advice is the same but the log line should not look like a bug.
    log.info(`github-auth: refresh token for ${userId} has expired; session dropped`);
    clearSession(userId);
    return { ok: false, reason: "session-expired" };
  }

  log.info(`github-auth: access token for ${userId} is stale; refreshing`);

  try {
    const refreshed = await exchange(
      {
        grant_type: "refresh_token",
        refresh_token: session.refreshToken as string,
      },
      log
    );

    const updated = toSession(refreshed, session.login, session.createdAt);
    setSession(userId, updated);
    return { ok: true, token: updated.accessToken, login: updated.login };
  } catch (err) {
    log.error(`github-auth: refresh failed for ${userId}`, describeError(err), err);
    clearSession(userId);
    return { ok: false, reason: "refresh-failed" };
  }
}

/**
 * Completes the OAuth flow. Called by the callback route below.
 *
 * Returns the user the session was established for, so the route can log it,
 * or null when the state was unknown, replayed, or expired.
 */
export async function completeSignIn(
  code: string,
  state: string,
  log: ILogger
): Promise<{ userId: string; login?: string } | null> {
  const pending = redeemPendingState(state);
  if (!pending) {
    // Not necessarily an attack -- a bookmarked callback URL, a double-clicked
    // link, or someone who took longer than ten minutes will all land here.
    log.warn("github-auth: OAuth callback presented an unknown, replayed, or expired state");
    return null;
  }

  const tokens = await exchange(
    {
      code,
      redirect_uri: callbackUrl(),
    },
    log
  );

  // Best-effort: knowing the login lets messages say *which* GitHub account is
  // connected, which matters when somebody has a work and a personal account.
  // Failing to get it must not cost the user a working session.
  let login: string | undefined;
  try {
    login = await fetchLogin(tokens.access_token, log);
  } catch (err) {
    log.warn("github-auth: could not read the GitHub login for a new session", describeError(err));
  }

  const session = toSession(tokens, login, Date.now());
  setSession(pending.userId, session);

  log.info(
    `github-auth: session established for ${pending.userId}` +
      (login ? ` (github login ${login})` : "") +
      (session.expiresAt ? `, access token expires ${new Date(session.expiresAt).toISOString()}` : "")
  );

  return { userId: pending.userId, login };
}

/** Forgets a user's GitHub session locally. Does NOT revoke it at GitHub. */
export function signOut(userId: string): boolean {
  const had = Boolean(getSession(userId));
  clearSession(userId);
  return had;
}

interface TokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in?: number;
  refresh_token_expires_in?: number;
  token_type?: string;
  error?: string;
  error_description?: string;
}

/**
 * POSTs to GitHub's token endpoint, for both the initial code exchange and
 * refreshes -- the endpoint and error handling are identical, only the grant
 * differs.
 *
 * NOTE: GitHub answers a *failed* exchange with HTTP 200 and an `error` field
 * in the body, not a 4xx. Checking the status code alone would treat an expired
 * code as a successful sign-in with an undefined token, so the body check below
 * is the real error path and must not be removed.
 */
async function exchange(
  grant: Record<string, string>,
  log: ILogger
): Promise<TokenResponse> {
  const body = new URLSearchParams({
    client_id: config.github.clientId,
    client_secret: config.github.clientSecret,
    ...grant,
  });

  const response = await fetch(GITHUB_TOKEN_URL, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: body.toString(),
    signal: AbortSignal.timeout(TOKEN_REQUEST_TIMEOUT_MS),
  });

  if (!response.ok) {
    throw new Error(`GitHub token endpoint returned HTTP ${response.status}`);
  }

  const data = (await response.json()) as TokenResponse;

  if (data.error || !data.access_token) {
    // error_description is GitHub's own prose and is safe to log; it names the
    // actual cause (bad_verification_code, incorrect_client_credentials, ...)
    // which is the difference between "the user was slow" and "the App is
    // misconfigured".
    log.error(
      `github-auth: token exchange rejected -- ${data.error ?? "no access_token in response"}` +
        (data.error_description ? `: ${data.error_description}` : "")
    );
    throw new Error(data.error ?? "GitHub returned no access token");
  }

  return data;
}

async function fetchLogin(accessToken: string, log: ILogger): Promise<string | undefined> {
  const response = await fetch(GITHUB_USER_URL, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
    },
    signal: AbortSignal.timeout(TOKEN_REQUEST_TIMEOUT_MS),
  });

  if (!response.ok) {
    log.warn(`github-auth: /user returned HTTP ${response.status}`);
    return undefined;
  }

  const data = (await response.json()) as { login?: string };
  return data.login;
}

/** Turns a token response into a stored session, converting relative expiries to absolute. */
function toSession(
  tokens: TokenResponse,
  login: string | undefined,
  createdAt: number
): GitHubSession {
  const now = Date.now();
  return {
    accessToken: tokens.access_token,
    refreshToken: tokens.refresh_token,
    expiresAt: tokens.expires_in ? now + tokens.expires_in * 1000 : undefined,
    refreshTokenExpiresAt: tokens.refresh_token_expires_in
      ? now + tokens.refresh_token_expires_in * 1000
      : undefined,
    login,
    createdAt,
  };
}

/**
 * The minimum of Express's request/response shapes that this route needs.
 *
 * Declared locally rather than imported from `express`, which is only a
 * transitive dependency here and ships no types of its own. Keeping the surface
 * this small also means the route body stays readable and would survive a move
 * to a different HTTP framework.
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
 * /api/messages.
 *
 * DELIBERATELY OUTSIDE THE TEAMS AUTH MIDDLEWARE. /api/messages is registered
 * through HttpServer, which validates a Bot Framework JWT on every request.
 * This route is registered straight onto the Express adapter instead, because
 * the caller is a browser following a redirect from github.com and has no Teams
 * token to present.
 *
 * That is safe for exactly one reason: the route trusts nothing except a
 * `state` value this process generated, stored server-side, and will accept
 * only once within ten minutes. Everything about who the session belongs to
 * comes from that server-side entry -- never from the query string. If the
 * state does not redeem, the request is refused before a single call is made to
 * GitHub.
 */
export function registerGitHubOAuthRoutes(router: RouteRegistrar, log: ILogger): void {
  if (!isGitHubConfigured()) {
    log.warn(
      "github-auth: no GitHub App configured (GITHUB_CLIENT_ID / GITHUB_CLIENT_SECRET / a " +
        "resolvable public origin), so the OAuth callback route was not mounted and " +
        "search_github will report itself unavailable."
    );
    return;
  }

  router.get(GITHUB_CALLBACK_PATH, async (req: CallbackRequest, res: CallbackResponse) => {
    const code = typeof req.query.code === "string" ? req.query.code : "";
    const state = typeof req.query.state === "string" ? req.query.state : "";
    const oauthError = typeof req.query.error === "string" ? req.query.error : "";

    // The user pressed Cancel on GitHub's authorization screen, or GitHub
    // refused. Not our failure, and it should not read like one.
    if (oauthError) {
      log.info(`github-auth: user did not complete authorization (${oauthError})`);
      res
        .status(200)
        .type("text/html")
        .send(page("GitHub was not connected", "You cancelled the sign-in, so nothing changed. You can close this tab and try again from Teams whenever you like."));
      return;
    }

    if (!code || !state) {
      res
        .status(400)
        .type("text/html")
        .send(page("Something went wrong", "That sign-in link was incomplete. Please start again from Teams."));
      return;
    }

    try {
      const result = await completeSignIn(code, state, log);

      if (!result) {
        res
          .status(400)
          .type("text/html")
          .send(page("That link has expired", "Sign-in links are only good for a few minutes and can be used once. Please ask Knowva again in Teams to get a fresh one."));
        return;
      }

      res
        .status(200)
        .type("text/html")
        .send(
          page(
            "GitHub connected",
            (result.login
              ? `Knowva is now connected to GitHub as ${escapeHtml(result.login)}.`
              : "Knowva is now connected to GitHub.") +
              " You can close this tab and go back to Teams."
          )
        );
    } catch (err) {
      log.error("github-auth: OAuth callback failed", describeError(err), err);
      res
        .status(500)
        .type("text/html")
        .send(page("Something went wrong", "Knowva couldn't finish connecting to GitHub. That's a problem on our side -- please try again from Teams in a moment."));
    }
  });

  log.info(`github-auth: OAuth callback mounted at ${callbackUrl()}`);
}

/**
 * A minimal self-contained results page.
 *
 * Every interpolated value is escaped at the call site. The only dynamic value
 * that ever reaches here is a GitHub login, but a page that renders in a user's
 * browser is not the place to rely on a value happening to be safe today.
 */
function page(heading: string, body: string): string {
  return (
    "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">" +
    "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">" +
    `<title>${escapeHtml(heading)} - Knowva</title>` +
    "<style>body{font-family:system-ui,-apple-system,Segoe UI,sans-serif;margin:0;" +
    "display:grid;place-items:center;min-height:100vh;background:#faf9f8;color:#201f1e}" +
    "main{max-width:32rem;padding:2rem;text-align:center}" +
    "h1{font-size:1.35rem;margin:0 0 .75rem}p{margin:0;line-height:1.5;color:#484644}</style>" +
    `</head><body><main><h1>${escapeHtml(heading)}</h1><p>${body}</p></main></body></html>`
  );
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
