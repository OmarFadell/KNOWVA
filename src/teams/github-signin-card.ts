import { MessageActivityInput } from "@microsoft/teams.api";
import { AdaptiveCard, OpenUrlAction, TextBlock } from "@microsoft/teams.cards";

/**
 * The GitHub sign-in card.
 *
 * A PLAIN ADAPTIVE CARD WITH A LINK -- deliberately not a Bot Framework OAuth
 * card, and not `ctx.signin()`. Those are bound to the Azure Bot Service OAuth
 * connection named `graph`, which exists for Microsoft Graph and knows nothing
 * about GitHub. Handing Teams an OAuth card here would start a token exchange
 * against the wrong identity provider. See the header of src/auth/github.ts.
 *
 * The card's job is to carry a URL the model must never touch: it is
 * single-use, expires in ten minutes, and is bound to one user's pending
 * sign-in. That is why it travels as a signal from the tool to the response
 * layer rather than as text through the model (see AgentToolSignal in
 * src/agent/loop.ts).
 *
 * The wording does two things beyond offering the button. It says this is
 * GitHub and not Teams, because a user who has already signed into Teams
 * silently will otherwise reasonably wonder why they are being asked again.
 * And it says what the connection can see, because "authorize Knowva" on
 * GitHub's own screen lists permissions in GitHub's terms, not in terms of what
 * Knowva will do with them.
 */
export function gitHubSignInCard(signInUrl: string): MessageActivityInput {
  const card = new AdaptiveCard(
    new TextBlock("Connect GitHub", { size: "Large", weight: "Bolder", wrap: true }),
    new TextBlock(
      "To search GitHub I need you to connect your own GitHub account. You'll only see " +
        "repositories you already have access to -- I can't see anything you can't.",
      { wrap: true }
    ),
    new TextBlock(
      "This is separate from your Microsoft Teams sign-in, and it's read-only: I can read " +
        "code, issues and pull requests, and I can't change anything.",
      { wrap: true, isSubtle: true, size: "Small" }
    )
  ).withActions(new OpenUrlAction(signInUrl, { title: "Connect GitHub" }));

  // The fallback text matters more than usual: it is what shows in the Teams
  // notification and in clients that cannot render a card at all, where a bare
  // "Knowva sent a card" would leave the user with no idea what to do.
  return new MessageActivityInput("Connect your GitHub account to search GitHub.").addCard(
    "adaptive",
    card
  );
}
