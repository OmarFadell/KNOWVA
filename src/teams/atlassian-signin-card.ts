import { MessageActivityInput } from "@microsoft/teams.api";
import { AdaptiveCard, OpenUrlAction, TextBlock } from "@microsoft/teams.cards";

/**
 * The Jira sign-in card. The Atlassian twin of
 * src/teams/github-signin-card.ts, and the mechanics there apply unchanged: a
 * PLAIN ADAPTIVE CARD WITH A LINK, deliberately not a Bot Framework OAuth card
 * and not `ctx.signin()`, because those are bound to the `graph` OAuth
 * connection, which exists for Microsoft Graph. The card carries a URL the
 * model must never touch -- single-use, ten-minute expiry, bound to one user's
 * pending sign-in -- which is why it travels as a signal from the tool to the
 * response layer rather than as text through the model.
 *
 * ===========================================================================
 * THE WORDING IS DOING REAL WORK HERE, AND IT HAD TO BE REWRITTEN.
 *
 * THIS CARD USED TO SAY "I don't keep any access to your Jira afterwards."
 * That was true when the Atlassian flow resolved an account id and discarded
 * the token. IT IS NOW FALSE. Confluence search runs as the user, so Knowva
 * retains a refreshable session and holds real, ongoing read access for as long
 * as the grant stands.
 *
 * Leaving the old wording in place would have been the worst kind of stale
 * copy: a specific, reassuring, incorrect claim about what Knowva keeps, made
 * at the exact moment the user is deciding whether to grant it. So the card now
 * says plainly that the connection stays, what it is used for, and how to end
 * it.
 *
 * It still names the two uses separately, because they really are different and
 * the consent screen will show scopes for both: finding Jira items that mention
 * them (which needs only their account details) and searching Confluence (which
 * needs ongoing read access). A user who reads the Atlassian grant screen sees
 * Confluence scopes, and this card should have prepared them for that.
 * ===========================================================================
 */
export function atlassianSignInCard(signInUrl: string): MessageActivityInput {
  const card = new AdaptiveCard(
    new TextBlock("Connect Atlassian", { size: "Large", weight: "Bolder", wrap: true }),
    new TextBlock(
      "Connecting lets me search Confluence as you, and work out which Jira items mention " +
        "you. You'll only ever see pages you already have access to — Confluence applies your " +
        "own permissions, so I can't see anything you can't.",
      { wrap: true }
    ),
    new TextBlock(
      "I keep this connection until you end it, so I can search without asking you to sign in " +
        "again. It's read-only — I can't change anything. Say **/jira-signout** to disconnect, " +
        "or remove Knowva under Connected apps in your Atlassian account settings to revoke it " +
        "fully.",
      { wrap: true, isSubtle: true, size: "Small" }
    ),
    new TextBlock("This is separate from your Microsoft Teams sign-in and from GitHub.", {
      wrap: true,
      isSubtle: true,
      size: "Small",
    })
  ).withActions(new OpenUrlAction(signInUrl, { title: "Connect Atlassian" }));

  // The fallback text matters more than usual: it is what shows in the Teams
  // notification and in clients that cannot render a card at all, where a bare
  // "Knowva sent a card" would leave the user with no idea what to do.
  return new MessageActivityInput(
    "Connect your Atlassian account so I can search Confluence and find Jira items that mention you."
  ).addCard("adaptive", card);
}
