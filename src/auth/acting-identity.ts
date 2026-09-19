import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import { describeError } from "../errors";
import { getMe } from "../graph/me";

/**
 * Answers one question, once per turn: WHOSE identity is this tool acting on
 * behalf of?
 *
 * ===========================================================================
 * WHY THIS EXISTS SEPARATELY FROM THE TEAMS ACTIVITY
 *
 * The obvious answer is "whoever sent the message", read off
 * `activity.from.aadObjectId`. That is usually right and it is not good enough
 * on its own, because it is a claim made by the routing layer about who is
 * talking, not proof of whose credentials the tool is about to spend.
 *
 * So the acting identity is resolved from the *token itself* -- `/me` -> `id`,
 * an Entra object id -- and the id Teams put on the activity is treated as a
 * second opinion. When both are available they must agree; when they disagree,
 * something is wrong enough that no per-user credential should be spent.
 *
 * WHY THAT MATTERS MORE FOR GITHUB THAN IT DID FOR EMAIL. The email tools have
 * defence in depth: even if the opt-out were checked against the wrong person,
 * `/me/messages` still runs on the asking user's delegated Graph token and can
 * only ever open that one mailbox. Graph would refuse to do the wrong thing.
 *
 * GitHub has no such backstop. The session lookup in src/auth/github.ts is the
 * *only* thing standing between "Alice asked a question" and "we sent Bob's
 * GitHub token to GitHub". GitHub cannot tell: it receives a valid bearer token
 * and answers for whoever it belongs to. That is the confused-deputy problem
 * the per-user OAuth design exists to avoid, and it is defeated entirely by one
 * wrong lookup key. Hence: resolve from the token, cross-check the activity,
 * and fail closed on any doubt.
 * ===========================================================================
 *
 * FAILS CLOSED. Every failure mode -- lookup error, missing id, mismatch --
 * returns a reason rather than a best guess. A caller that cannot establish who
 * it is acting for must refuse to act.
 *
 * NOTE ON DUPLICATION: src/tools/email-access.ts contains an equivalent
 * resolver, written before this one and deliberately left alone here (this
 * milestone was scoped not to modify the existing tools). Two copies of an
 * identity check is exactly the kind of thing that drifts apart, and the email
 * one should be folded into this module the next time it is touched. Behaviour
 * is intentionally identical today, minus the mismatch check, which is new.
 */

/**
 * Deliberately one shape with optional fields rather than a discriminated
 * union. This project compiles without `strictNullChecks`, which switches off
 * narrowing on `ok: true | false`, so a union here would force every caller
 * into a type assertion -- and an assertion around a security check is exactly
 * the wrong place to be silencing the compiler. Same reasoning as the explicit
 * `isToolError` guard in src/tools/email-access.ts.
 */
export interface ActingIdentity {
  /** True only when `userId` is an identity the caller may act as. */
  ok: boolean;
  /** Set when ok. The Entra object id to key per-user credentials on. */
  userId?: string;
  /** Set when not ok. Why the identity could not be trusted. */
  reason?: "unresolved" | "mismatch";
}

export interface ActingIdentityOptions {
  /** This turn's delegated Graph client -- the asking user's token. */
  graph: GraphClient;
  /** `activity.from.aadObjectId`, when Teams supplied one. */
  claimedUserId?: string;
  /**
   * The acting id if the caller already resolved it via /me this turn. Purely a
   * round-trip saving; when absent this module resolves it rather than trusting
   * `claimedUserId` on its own.
   */
  resolvedUserId?: string;
  log: ILogger;
}

export interface ActingIdentityResolver {
  /** Memoised for the turn: identity cannot change mid-turn, and nor can the answer. */
  resolve(): Promise<ActingIdentity>;
}

export function createActingIdentityResolver(
  options: ActingIdentityOptions
): ActingIdentityResolver {
  const { graph, claimedUserId, log } = options;

  let cached: ActingIdentity | undefined;

  return {
    async resolve(): Promise<ActingIdentity> {
      if (cached) return cached;

      let actingUserId = options.resolvedUserId;

      if (!actingUserId) {
        try {
          const me = await getMe(graph);
          actingUserId = me.id;
          if (!me.id) {
            log.error("acting-identity: /me returned no id");
          }
        } catch (err) {
          log.error("acting-identity: /me lookup failed", describeError(err), err);
        }
      }

      if (!actingUserId) {
        cached = { ok: false, reason: "unresolved" };
        return cached;
      }

      // Both known and disagreeing is the case worth refusing loudly. It should
      // be impossible -- the token was minted for the sender of the activity --
      // so if it ever happens, something about how identity reaches this turn
      // has changed, and continuing would mean spending one person's credential
      // on another person's question.
      if (claimedUserId && !idsMatch(actingUserId, claimedUserId)) {
        log.error(
          "acting-identity: the token's user and the activity's user disagree -- refusing. " +
            `token=${actingUserId} activity=${claimedUserId}`
        );
        cached = { ok: false, reason: "mismatch" };
        return cached;
      }

      cached = { ok: true, userId: actingUserId };
      return cached;
    },
  };
}

function idsMatch(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}
