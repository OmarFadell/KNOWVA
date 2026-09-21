import type { MessageActivityInput } from "@microsoft/teams.api";
import type { AuthProvider } from "../agent/loop";
import { buildSignInUrl as buildGitHubSignInUrl } from "../auth/github";
import { buildSignInUrl as buildAtlassianSignInUrl } from "../auth/atlassian";
import { gitHubSignInCard } from "./github-signin-card";
import { atlassianSignInCard } from "./atlassian-signin-card";

/**
 * The registry that turns a `needs-auth` signal into an actual sign-in card.
 *
 * ---------------------------------------------------------------------------
 * WHY THIS FILE EXISTS, AND WHAT IT REPLACED.
 *
 * Before Atlassian there was exactly one provider, and app.ts said so in three
 * separate places: a hardcoded `signal.provider !== "github"` test, a direct
 * import of the GitHub buildSignInUrl, and a direct import of the GitHub card.
 * That was the right amount of structure for one provider.
 *
 * Adding a second would have meant either an if/else ladder in the response
 * layer or a copy of sendAuthCards per provider -- and the response layer is
 * the wrong place for "which OAuth app does this provider use", because it has
 * no other reason to know that any of them exist.
 *
 * So the knowledge lives here, one entry per provider, and sendAuthCards in
 * app.ts becomes provider-agnostic: look the signal up, build the URL, send the
 * card. A third provider is one entry in this table and touches nothing else.
 *
 * NOTE THE ONE PIECE THAT DOES NOT GENERALISE: `buildUrl` returns null when
 * that provider is not configured. It has to stay per-provider, because
 * "configured" means different things per provider (a GitHub App's client
 * credentials, an Atlassian OAuth integration's) and each module already owns
 * that judgement. The registry deliberately does not try to unify it.
 * ---------------------------------------------------------------------------
 */
export interface SignInOffer {
  /**
   * The authorize URL for this user, or null when the provider is not
   * configured on this deployment. The URL is single-use and bound to the
   * user, which is why it is built at send time rather than passed around.
   */
  buildUrl(userId: string): string | null;
  /** The Adaptive Card carrying that URL. */
  card(signInUrl: string): MessageActivityInput;
}

export const signInOffers: Record<AuthProvider, SignInOffer> = {
  github: {
    buildUrl: buildGitHubSignInUrl,
    card: gitHubSignInCard,
  },
  atlassian: {
    buildUrl: buildAtlassianSignInUrl,
    card: atlassianSignInCard,
  },
};
