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
    "from the chat to stop it for good -- once I'm removed I lose access entirely."
  );
}
