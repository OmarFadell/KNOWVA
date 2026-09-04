import { Client as GraphClient } from "@microsoft/teams.graph";

/**
 * Builds a Graph client bound to a specific user token.
 *
 * Needed when a token arrives part-way through a turn: the activity context's
 * own `userGraph` closes over the eager token fetch made at the start of the
 * turn, so it stays unauthenticated even after `signin()` returns a token.
 */
export function graphClientForToken(token: string): GraphClient {
  return new GraphClient({ token });
}
