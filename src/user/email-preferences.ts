/**
 * Per-user opt-out for the Outlook email tools.
 *
 * WHY THIS IS PER *USER* AND NOT PER CONVERSATION
 *
 * Every other switch in Knowva is scoped to a conversation: /pause quiets one
 * chat, /reset forgets one thread. This one cannot be, because the thing being
 * protected is not a conversation -- it is a mailbox. Somebody who does not want
 * Knowva reading their mail does not want it reading their mail in a group chat
 * either, and would have no way to know which chats to visit. So the flag
 * follows the person, and one /disable-email anywhere disables mail search
 * everywhere for them.
 *
 * KEYED ON THE ENTRA OBJECT ID, ON PURPOSE
 *
 * Not on `activity.from.id`, which is a Teams-surface identifier that varies by
 * channel and tells you nothing about which mailbox a Graph token will open.
 * The object id is what `/me` returns as `id` and what Teams puts on
 * `activity.from.aadObjectId`, so the id written by /disable-email and the id
 * checked before a Graph mail call are the same value, arrived at from two
 * different directions. That is what makes the check meaningful rather than
 * merely plausible.
 *
 * ---------------------------------------------------------------------------
 * IN-MEMORY ONLY -- AND HERE THAT IS WORSE THAN IT IS ELSEWHERE.
 *
 * Like src/llm/history.ts and the paused-conversation set in
 * src/teams/group-chat.ts, this is a module-level Map that dies with the
 * process: every deploy, every idle recycle on App Service, and it is
 * per-instance so a second worker never sees it.
 *
 * For history that is a lost convenience. For this it is a privacy regression:
 * a user's opt-out silently lapses and Knowva starts reading their mailbox
 * again, with nothing to tell them it happened. This store therefore fails
 * OPEN on restart, which is the wrong direction for a control of this kind, and
 * it is the single strongest argument for giving Knowva durable storage. When
 * that lands, this file is the first thing that should move to it -- ahead of
 * conversation history, which merely degrades.
 *
 * Until then the honest mitigation is to say so: the disable confirmation
 * message tells the user the opt-out may not survive a restart, rather than
 * promising a permanence this cannot deliver.
 * ---------------------------------------------------------------------------
 */

/** Entra object ids (lowercased) of users who have run /disable-email. */
const emailDisabled = new Set<string>();

/**
 * Object ids are GUIDs whose casing is not guaranteed stable between the Teams
 * activity and a Graph response, so everything is compared lowercased.
 */
function normalize(userId: string): string {
  return userId.trim().toLowerCase();
}

export function disableEmailForUser(userId: string): void {
  emailDisabled.add(normalize(userId));
}

export function enableEmailForUser(userId: string): void {
  emailDisabled.delete(normalize(userId));
}

export function isEmailDisabledForUser(userId: string | undefined | null): boolean {
  if (!userId) return false;
  return emailDisabled.has(normalize(userId));
}

/**
 * True if *any* of the supplied identities has opted out.
 *
 * Takes a list rather than one id so a caller can pass every identity involved
 * in a turn -- the person chatting and the person whose token the call will
 * actually run as -- and have the tool refuse if either has opted out. In the
 * current architecture those are always the same person (see
 * src/tools/email-access.ts), so this reads as belt and braces; the point is
 * that if a future on-behalf-of path ever makes them diverge, this fails closed
 * rather than quietly picking whichever one happens to be checked.
 */
export function isEmailDisabledForAny(userIds: (string | undefined | null)[]): boolean {
  return userIds.some((id) => isEmailDisabledForUser(id));
}

/**
 * What a tool hands back to the model when the user has opted out.
 *
 * Deliberately not an error, and deliberately not generic. "Something went
 * wrong" would invite the model to apologise for a fault and suggest retrying;
 * this is a working feature behaving exactly as the user asked, and the reply
 * should say so and name the way back.
 */
export function emailDisabledToolMessage(): string {
  return (
    "Email search is switched off for this user: they have run /disable-email, so Knowva will " +
    "not read their mailbox. This is not an error and nothing needs retrying. Tell them plainly " +
    "that you can't search their email because they turned it off, and that **/enable-email** " +
    "turns it back on. Answer the rest of their question from other sources if you can, and do " +
    "NOT guess at what an email might have said."
  );
}
