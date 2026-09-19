import type { ILogger } from "@microsoft/teams.common";
import type { AgentToolResult } from "../agent/loop";
import type { ActingIdentityResolver } from "../auth/acting-identity";
import { getValidAccessToken, isGitHubConfigured, type GitHubTokenFailure } from "../auth/github";
import type { McpFailure } from "../github/mcp-client";

/**
 * The gate every GitHub tool passes through, shared by search_github and
 * get_repo_content.
 *
 * Same role as src/tools/email-access.ts plays for the two email tools, and it
 * exists for the same reason: two tools that answer "may this user do this?"
 * from two copies of the logic will eventually answer it differently, and the
 * one that drifts will be the one nobody reads.
 *
 * ===========================================================================
 * THE PROPERTY THIS FILE PROTECTS
 *
 * Knowva holds a GitHub token for every user who has signed in, and this is
 * where one gets chosen. There is no second line of defence: GitHub receives a
 * valid bearer token and answers for whoever it belongs to, so picking the
 * wrong session is undetectable from the far end. Both tools therefore resolve
 * the acting user through src/auth/acting-identity.ts -- which reads the
 * identity out of the Graph token and cross-checks it against the id Teams put
 * on the activity -- and refuse on any doubt, including the two disagreeing.
 *
 * In a group chat that is the difference that matters: the turn runs on the
 * asker's token, so the asker's GitHub session is the only one in play, no
 * matter who else is in the room or who signed in most recently.
 * ===========================================================================
 */

export interface GitHubTokenLookup {
  /** True only when `token` is this user's live GitHub access token. */
  ok: boolean;
  token?: string;
  /** The GitHub login the token belongs to, when known. */
  login?: string;
  /** Set when not ok: a ready-made result for the tool to return as-is. */
  refusal?: AgentToolResult;
}

/**
 * Resolves the acting user, then their GitHub token.
 *
 * Returns a ready-made `refusal` rather than a reason code, because every
 * caller would otherwise translate the same four failures into the same four
 * messages. A tool just returns it.
 */
export async function resolveGitHubToken(
  identity: ActingIdentityResolver,
  log: ILogger
): Promise<GitHubTokenLookup> {
  if (!isGitHubConfigured()) {
    return { ok: false, refusal: notConfiguredResult() };
  }

  // Nothing that spends a credential may run before this.
  const acting = await identity.resolve();
  if (!acting.ok) {
    log.warn(`github-access: refusing -- acting identity ${acting.reason}`);
    return { ok: false, refusal: identityRefusalResult(acting.reason) };
  }

  const token = await getValidAccessToken(acting.userId, log);
  if (!token.ok) {
    return { ok: false, refusal: needsAuthResult(token.reason) };
  }

  return { ok: true, token: token.token, login: token.login };
}

export function notConfiguredResult(): AgentToolResult {
  return {
    content:
      "GitHub isn't set up on this deployment -- Knowva has no GitHub App configured. Tell the " +
      "user it's a setup problem on our side, not something they did.",
    isError: true,
  };
}

function identityRefusalResult(reason: string | undefined): AgentToolResult {
  return {
    content:
      reason === "mismatch"
        ? "Knowva could not safely establish whose GitHub account to act as, so it did nothing. " +
          "This is a safety stop. Tell the user something is wrong with their sign-in and to " +
          "sign out and back in."
        : "Knowva could not confirm who it would be acting as on GitHub, so it did nothing. " +
          "This is a safety stop, not a permissions problem. Tell the user to sign out and sign " +
          "back in, then try again.",
    isError: true,
  };
}

/**
 * The "sign in first" result: prose for the model, plus a signal that makes
 * app.ts attach a sign-in card (see AgentToolSignal in src/agent/loop.ts).
 *
 * The reasons are kept apart because they read differently to somebody who has
 * been using this successfully for hours. "Connect GitHub" is right for a user
 * who never has; "your connection expired" is right for one whose 8-hour token
 * ran out, and telling them to connect would imply their earlier sign-in never
 * happened.
 */
export function needsAuthResult(reason: GitHubTokenFailure | undefined): AgentToolResult {
  if (reason === "not-configured") return notConfiguredResult();

  // `reason` is always set when the token result is not ok; the fallback exists
  // so a future failure mode added without a message here degrades to the safe,
  // accurate one rather than to an empty explanation.
  const explanation =
    reason === "session-expired"
      ? "This user's GitHub connection has expired and cannot be renewed automatically, so " +
        "nothing was fetched."
      : reason === "refresh-failed"
        ? "This user's GitHub connection could not be renewed, so nothing was fetched."
        : "This user hasn't connected their GitHub account yet, so nothing was fetched.";

  return {
    // No URL in here on purpose -- the card carries it.
    content:
      explanation +
      " Knowva is showing them a sign-in button, so do NOT paste a link or invent one. Tell " +
      "them briefly that you need them to connect GitHub first and that there's a button just " +
      "below your message, then stop. Say nothing about what you might have found -- you " +
      "fetched nothing. Note this is a GitHub sign-in, entirely separate from their Microsoft " +
      "Teams sign-in.",
    signal: { kind: "needs-auth", provider: "github" },
  };
}

/**
 * Transport- and protocol-level MCP failures, shared by both tools.
 *
 * `what` names the operation in the model's terms ("that code search", "that
 * file") so one set of messages can serve both without reading oddly.
 * Tool-level errors are NOT handled here: those carry GitHub's own explanation
 * and mean different things per tool, so each tool classifies its own.
 */
export function mcpFailureResult(
  reason: McpFailure | undefined,
  detail: string | undefined,
  what: string
): AgentToolResult {
  switch (reason) {
    case "unauthorized":
      // GitHub rejected a token we believed was live -- most likely the user
      // revoked the app, or an admin removed the installation.
      return {
        content:
          "GitHub rejected this user's credentials. Their connection may have been revoked on " +
          "GitHub's side. Tell them to reconnect by asking again -- Knowva will offer a fresh " +
          "sign-in button -- and note this is the GitHub connection, not their Teams sign-in.",
        isError: true,
        signal: { kind: "needs-auth", provider: "github" },
      };

    case "forbidden":
      return {
        content:
          `GitHub refused ${what}. The most likely reason is that the GitHub App isn't installed ` +
          "on the organization or repository in question, or this user's account has no access " +
          "to it. Tell the user plainly, and suggest they check with whoever administers that " +
          "organization. Do not guess at what the result would have been.",
        isError: true,
      };

    case "rate-limited":
      return {
        content:
          "GitHub is rate-limiting this account. Tell the user to try again in a few minutes.",
        isError: true,
      };

    case "transport-error":
    default:
      return {
        content:
          `Knowva couldn't reach GitHub for ${what}. Tell the user it's a problem on our side, ` +
          "not theirs, and to try again shortly. Do not guess at what it would have returned." +
          (detail ? ` (Technical detail, not for the user: ${detail})` : ""),
        isError: true,
      };
  }
}
