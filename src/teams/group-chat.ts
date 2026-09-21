import type { IActivity } from "@microsoft/teams.api";

/**
 * Group-chat behaviour: what Knowva reads, and what it answers.
 *
 * THE DISTINCTION THIS FILE EXISTS TO ENFORCE
 *
 * Knowva receives every message in a group chat it is installed in, whether or
 * not it was @mentioned -- that is what the `groupChat` bot scope means, and it
 * is the point: context is what makes a later answer useful.
 *
 * Reading is not a licence to talk. A bot that replies to everything in a busy
 * group chat is a bot people uninstall within the hour. So the two decisions are
 * kept deliberately separate and named separately:
 *
 *   shouldIngest(activity)  -- may we record this for context?   (almost always)
 *   shouldRespond(activity) -- should we say something back?      (only if addressed)
 *
 * Every caller must ask both. Neither implies the other.
 */

/** Teams marks group chats and channel conversations as groups; personal chats are 1:1 with the bot. */
export function isGroupConversation(activity: IActivity): boolean {
  const conversation = activity.conversation;
  return conversation.isGroup === true || conversation.conversationType === "groupChat";
}

/**
 * Was the bot actually addressed?
 *
 * True when the bot is @mentioned, or when the conversation is personal (a 1:1
 * chat with Knowva is addressed to Knowva by definition -- there is nobody else
 * in it). Everything else in a group chat is other people talking to each other.
 *
 * A reply to one of Knowva's own messages does NOT count. Teams does not
 * reliably surface that as a mention, and guessing wrong means interjecting --
 * the exact failure this gate exists to prevent. Better to need one @mention
 * than to become the bot that butts in.
 */
export function isAddressedToBot(activity: IActivity): boolean {
  if (!isGroupConversation(activity)) return true;

  const botId = activity.recipient?.id;
  if (!botId) return false;

  return (activity.entities ?? []).some(
    (entity) =>
      entity.type === "mention" &&
      // Teams prefixes bot ids inconsistently across surfaces (e.g. "28:<id>"),
      // so compare on the tail rather than requiring an exact match.
      idsMatch((entity as { mentioned?: { id?: string } }).mentioned?.id, botId)
  );
}

function idsMatch(a: string | undefined, b: string | undefined): boolean {
  if (!a || !b) return false;
  if (a === b) return true;
  const tail = (id: string) => id.slice(id.lastIndexOf(":") + 1);
  return tail(a) === tail(b);
}

/**
 * Conversations where someone ran /pause.
 *
 * A SOFT switch, and the welcome message says so. The real off switch is
 * uninstalling the app, which Teams enforces by revoking the RSC grant and
 * removing Knowva from the roster -- after that it cannot read the chat even if
 * this process wanted to. This set is only our own code choosing to be quiet,
 * and like every other in-memory store here it does not survive a restart.
 * A pause that silently lapses on deploy is the reason uninstall is presented
 * as the answer for anyone who actually needs Knowva gone.
 */
const pausedConversations = new Set<string>();

export function pauseConversation(conversationId: string): void {
  pausedConversations.add(conversationId);
}

export function resumeConversation(conversationId: string): void {
  pausedConversations.delete(conversationId);
}

export function isPaused(conversationId: string): boolean {
  return pausedConversations.has(conversationId);
}

/**
 * Should this activity be recorded for context?
 *
 * Yes for ordinary human messages in a chat that is not paused. System event
 * messages ("X added Knowva to the chat", renames, calls started) are excluded:
 * Teams delivers them as conversationUpdate activities rather than messages, and
 * Graph tags their stored form `systemEventMessage` -- src/graph/chat-messages.ts
 * drops those on the read side for the same reason. They are chat furniture, not
 * anything a person said.
 */
export function shouldIngest(activity: IActivity, text: string): boolean {
  if (activity.type !== "message") return false;
  if (isPaused(activity.conversation.id)) return false;
  if (!text.trim()) return false;
  return true;
}

/**
 * Should Knowva reply?
 *
 * Only when addressed, and never while paused. Note the asymmetry with
 * shouldIngest: a message can be worth remembering and still not be ours to
 * answer. That is the normal case in a group chat, not an edge case.
 */
export function shouldRespond(activity: IActivity): boolean {
  if (isPaused(activity.conversation.id)) return false;
  return isAddressedToBot(activity);
}

/**
 * What Knowva says when it is added to a group chat.
 *
 * Installation IS the consent -- there is no extra button, and Teams has already
 * posted its own "X added Knowva" system message naming who did it. So this is
 * not a consent request; it is a disclosure, and it has one job: make sure
 * nobody in the chat is surprised later about what Knowva can see.
 *
 * Four facts, one short paragraph, uninstall named as the real off switch.
 * Anything longer gets skimmed, and a disclosure nobody reads discloses nothing.
 */
export function groupChatWelcomeMessage(): string {
  return (
    "Hi -- I'm Knowva. Now that I'm in this chat I can read the messages here, and anyone " +
    "in the chat can @mention me to ask a question. I won't chime in unless I'm mentioned. " +
    "When I use something from this conversation I'll say who said it and when, and link to " +
    "the message. Use **/pause** if you want me to stop reading for a while, or remove me " +
    "from the chat to stop it for good -- once I'm removed I lose access entirely.\n\n" +
    "I can also search whoever asks me a question in **their own** Outlook inbox and sent " +
    "mail -- never yours, never anyone else's, and never drafts or deleted mail. If you'd " +
    "rather I never touched your mailbox, say **/disable-email** and I won't. **/help** " +
    "lists everything."
  );
}

/**
 * What Knowva says when somebody installs it for themselves, in a 1:1 chat.
 *
 * This did not exist before Outlook mail, and it exists now for one reason:
 * installing the app is what grants mail access, nobody clicks anything else,
 * and the store description says nothing about email. A user who never learns
 * Knowva can read their mailbox has not really opted into anything -- and an
 * opt-out only they can find by asking is not an opt-out.
 *
 * Same job as the group-chat disclosure, then: four facts and the way out, in
 * one short paragraph. Longer gets skimmed.
 */
export function personalWelcomeMessage(): string {
  return (
    "Hi -- I'm Knowva. Ask me anything and I'll answer from what I can reach: our SharePoint " +
    "documents, the Teams group chats I've been added to, and your Outlook email.\n\n" +
    "About email specifically, since it's your mailbox: I search your **Inbox and Sent Items " +
    "only** -- never drafts, never deleted mail, never anyone else's mailbox -- and I read " +
    "attachment filenames but never open the files. I always tell you who sent something and " +
    "when. If you'd rather I didn't read your mail at all, say **/disable-email** and I'll " +
    "stop; **/enable-email** turns it back on.\n\n" +
    "**/help** lists what I can do and every command."
  );
}

/**
 * The /help reply.
 *
 * Takes the caller's current email setting rather than describing both states,
 * so somebody who has switched email off is told that it *is* off -- the case
 * where a wrong answer would matter most, because it would leave them believing
 * their mail was still being read.
 */
export interface GitHubHelpState {
  /** False when no GitHub App is configured on this deployment. */
  configured: boolean;
  /** The connected GitHub login, when this user has one. */
  login?: string;
}

export function helpMessage(emailDisabled: boolean, github?: GitHubHelpState): string {
  const emailLine = emailDisabled
    ? "- **Email** -- currently **switched off** for you. Say **/enable-email** to turn it back on."
    : "- **Email** -- your Outlook Inbox and Sent Items. Never drafts, deleted mail, or anyone " +
      "else's mailbox. I read attachment filenames, never the files. Say **/disable-email** to " +
      "switch this off for you everywhere.";

  // Omitted entirely when no GitHub App is configured: offering a command that
  // cannot work is worse than not mentioning the feature.
  const gitHubLine = !github?.configured
    ? null
    : github.login
      ? `- **GitHub** -- connected as **${github.login}**. I search code, repos, issues and ` +
        "pull requests you already have access to. **/github-signout** disconnects."
      : "- **GitHub** -- not connected yet. Ask me something about a repo and I'll show you a " +
        "button to connect. You'll only ever see repositories your own GitHub account can.";

  return [
    "Here's what I can do:",
    "",
    "- **Documents** -- search our configured SharePoint site, and tell you when a document was " +
      "last changed and by whom.",
    "- **Conversations** -- search recent messages in the Teams group chats I've been added to " +
      "and that you're in. I always say who said what, when, and link to it.",
    emailLine,
    ...(gitHubLine ? [gitHubLine] : []),
    // Listed unconditionally, unlike the GitHub line above. That line is hidden
    // when no GitHub App is configured because offering a command that cannot
    // work is worse than not mentioning the feature -- but Atlassian config is
    // not plumbed through to this function, and describing Confluence search is
    // accurate for every deployment that has it. If a deployment without an
    // Atlassian app ever needs this hidden too, pass state in the way
    // GitHubHelpState already does rather than guessing here.
    "- **Confluence** -- search and read pages you already have access to. Confluence applies " +
      "your own permissions, so I can't see anything you can't. Connecting also lets me find " +
      "Jira test items that mention you. **/jira-signout** disconnects.",
    "",
    "Commands:",
    "",
    "- **/help** -- this message",
    "- **/disable-email** / **/enable-email** -- turn email search off or back on, for you, everywhere",
    ...(github?.configured
      ? ["- **/github-signout** -- disconnect my access to your GitHub account"]
      : []),
    "- **/jira-signout** -- disconnect my access to your Atlassian account (Confluence and Jira)",
    "- **/pause** / **/resume** -- stop or restart me reading this conversation",
    "- **/reset** -- forget what we've discussed here",
    "",
    "I can't send email, edit anything, read your calendar, or open attachments. I'm read-only.",
  ].join("\n");
}
