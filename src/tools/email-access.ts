import { Client as GraphClient } from "@microsoft/teams.graph";
import type { ILogger } from "@microsoft/teams.common";
import config from "../../config";
import type { SharedMailbox } from "../../config";
import { describeError } from "../errors";
import type { AgentToolResult } from "../agent/loop";
import { getMe } from "../graph/me";
import { mailboxPath } from "../graph/mail";
import { emailDisabledToolMessage, isEmailDisabledForAny } from "../user/email-preferences";

/**
 * The gate every Outlook mail read passes through, shared by search_emails and
 * get_email_content.
 *
 * Two jobs, both of which must happen before a single Graph mail request is
 * built: decide whether this user has opted out, and decide which mailbox is
 * even askable.
 *
 * ===========================================================================
 * WHOSE OPT-OUT IS CHECKED, AND WHY "THE PERSON CHATTING" IS THE RIGHT ANSWER
 * HERE (BUT WOULD NOT BE IN GENERAL)
 *
 * The milestone brief asked for the flag to be checked for the identity the
 * tool acts *on behalf of*, not merely the person typing -- so that a group
 * chat could not become a way to pull in somebody else's mail indirectly. In
 * this codebase those two identities are the same, and it is worth being
 * precise about why, because the reasoning does not generalise to every tool
 * here.
 *
 * Identity flows like this. Teams SSO exchanges a token for the user who sent
 * the activity -- `activity.from` -- and app.ts hands the resulting delegated
 * Graph client to the tool factories for that one turn. `/me/messages` on that
 * client therefore resolves to the sender of the message and to nobody else.
 * When Alice @mentions Knowva in a group chat, the token in play is Alice's;
 * Bob's mailbox is not reachable from it, and there is no impersonation, no
 * on-behalf-of exchange, and no app-only fallback anywhere on the mail path
 * (see the note at the top of src/graph/mail.ts, which forbids adding one).
 *
 * The tool that *does* juggle two identities is search_conversations: it lists
 * chats as the user but reads messages with the bot's own app-only token,
 * because RSC grants are application permissions. That divergence is real, and
 * it is exactly the shape that would break a naive "check the person chatting"
 * rule. Mail has no equivalent, because there is no delegated-vs-application
 * split to straddle -- an application Mail.Read would be tenant-wide, which is
 * why this feature does not have one.
 *
 * So: the check below resolves the acting identity from the *token itself*
 * (`/me` -> `id`, an Entra object id) rather than trusting the Teams-surface id
 * on the activity, and then refuses if EITHER that identity or the chatting
 * user's `aadObjectId` has opted out. Today those are the same value arrived at
 * two different ways, which makes the check a consistency assertion as much as
 * a permission test. If some future change ever does introduce an
 * on-behalf-of mail path, this fails closed instead of silently checking the
 * wrong person.
 *
 * FAIL CLOSED ON AN UNKNOWN IDENTITY. If the acting identity cannot be resolved
 * at all, the read is refused. An opt-out that can be bypassed by a failed
 * lookup is not an opt-out, and mail is not the place to guess.
 * ===========================================================================
 */

export interface EmailAccessOptions {
  /** This turn's delegated Graph client -- the asking user's token. */
  graph: GraphClient;
  /**
   * `activity.from.aadObjectId`: the Entra object id of whoever sent the
   * message, straight from Teams. Undefined on surfaces that omit it.
   */
  chattingUserId?: string;
  /**
   * The acting user's Entra object id, if app.ts already resolved it via /me
   * this turn. Passed in purely to avoid a second round trip; when it is
   * missing this module resolves it itself rather than skipping the check.
   */
  actingUserId?: string;
  log: ILogger;
}

/** A resolved mailbox the tools may read from. */
export interface ResolvedMailbox {
  /** Graph path prefix: "/me" or "/users/{address}". */
  path: string;
  /** Human-facing name for logs and for the model's citation, e.g. "your mailbox". */
  label: string;
}

export interface EmailAccess {
  /**
   * Returns a tool result to hand straight back when the read must not happen,
   * or null when it may proceed. Call this before building any Graph mail
   * request -- not after, and not around it.
   */
  guard(): Promise<AgentToolResult | null>;
  /**
   * Resolves the optional `mailbox` argument to a Graph path. Returns an error
   * result when the caller named something that is not configured.
   */
  resolveMailbox(requested: unknown): ResolvedMailbox | AgentToolResult;
}

export function createEmailAccess(options: EmailAccessOptions): EmailAccess {
  const { graph, chattingUserId, log } = options;

  // Memoised across both tools for one turn: Claude routinely calls
  // search_emails and then get_email_content in the same turn, and neither the
  // user's identity nor their opt-out can change in between.
  let actingUserId: string | undefined = options.actingUserId;
  let identityResolved = Boolean(options.actingUserId);
  let identityLookupFailed = false;

  async function resolveActingUser(): Promise<string | undefined> {
    if (identityResolved || identityLookupFailed) return actingUserId;

    try {
      const me = await getMe(graph);
      actingUserId = me.id;
      identityResolved = Boolean(me.id);
      if (!me.id) {
        log.error("email-access: /me returned no id; cannot establish the acting identity");
      }
    } catch (err) {
      identityLookupFailed = true;
      log.error(
        "email-access: could not resolve the acting user via /me",
        describeError(err),
        err
      );
    }

    return actingUserId;
  }

  return {
    async guard(): Promise<AgentToolResult | null> {
      const acting = await resolveActingUser();

      if (!acting) {
        // Fail closed. See the header: a check that a lookup failure can skip is
        // not a check. The message is deliberately about identity rather than
        // about mail, because that is what actually went wrong.
        return {
          content:
            "Knowva could not confirm whose mailbox it would be reading, so it did not read any " +
            "mail. This is a safety stop, not a permissions problem. Tell the user to sign out " +
            "and sign back in, then try again.",
          isError: true,
        };
      }

      if (isEmailDisabledForAny([acting, chattingUserId])) {
        log.info(`email-access: mail read refused -- user ${acting} has opted out`);
        return { content: emailDisabledToolMessage() };
      }

      return null;
    },

    resolveMailbox(requested: unknown): ResolvedMailbox | AgentToolResult {
      const name = typeof requested === "string" ? requested.trim() : "";

      // The default and overwhelmingly common case: the asking user's own
      // mailbox, reached with Mail.Read and no configuration at all.
      if (!name) {
        return { path: mailboxPath(), label: "your own mailbox" };
      }

      const match = findSharedMailbox(name, config.sharedMailboxes);
      if (match) {
        return {
          path: mailboxPath(match.address),
          label: `the ${match.name} shared mailbox (${match.address})`,
        };
      }

      // An unconfigured name is refused rather than guessed at. Passing the raw
      // string through to /users/{whatever} would turn a model hallucination
      // into an attempt to open an arbitrary colleague's mailbox -- which
      // Exchange would refuse, but which should never be asked in the first
      // place. The allow-list is the boundary; this is where it is enforced.
      const configured = config.sharedMailboxes;
      return {
        content:
          `There is no shared mailbox called "${name}". ` +
          (configured.length === 0
            ? "No shared mailboxes are configured on this deployment at all, so the only mailbox " +
              "you can search is the user's own -- call search_emails again without the mailbox " +
              "argument. Tell the user that is what you searched."
            : "The only shared mailboxes configured are: " +
              configured.map((mailbox) => `"${mailbox.name}"`).join(", ") +
              ". Either call search_emails again with one of those exact names, or omit the " +
              "mailbox argument to search the user's own mailbox. Do not invent a mailbox name."),
        isError: true,
      };
    },
  };
}

/** Narrows an unknown to the error branch of a resolveMailbox result. */
export function isToolError(value: ResolvedMailbox | AgentToolResult): value is AgentToolResult {
  return typeof (value as AgentToolResult).content === "string";
}

/**
 * Matches a caller-supplied mailbox name against the configured list.
 *
 * Accepts either the display name or the SMTP address, case-insensitively:
 * the model has seen both in the system prompt and in earlier tool results, and
 * refusing "projectalpha@contoso.com" because the config calls it "Project
 * Alpha" would be pedantry, not a boundary. Anything not on the list is still
 * refused.
 */
function findSharedMailbox(name: string, configured: SharedMailbox[]): SharedMailbox | undefined {
  const wanted = name.toLowerCase();
  return configured.find(
    (mailbox) =>
      mailbox.name.toLowerCase() === wanted || mailbox.address.toLowerCase() === wanted
  );
}

/**
 * Turns a failed Graph mail call into something the model can relay in plain
 * language. Same convention as the SharePoint tools: a user never sees a stack
 * trace, and a setup problem is never reported as if it were their fault.
 */
export function classifyMailError(
  err: unknown,
  mailbox: ResolvedMailbox,
  log: ILogger
): AgentToolResult {
  log.error(`email: Graph mail call failed for ${mailbox.path}`, describeError(err), err);

  const status = statusOf(err);
  const isShared = mailbox.path !== "/me";

  if (status === 401) {
    return {
      content:
        "The email search could not authenticate. Tell the user to sign out and sign back in, " +
        "then retry.",
      isError: true,
    };
  }

  if (status === 403) {
    // These two 403s have completely different fixes, and telling the user the
    // wrong one sends them to the wrong person. A shared mailbox they were
    // never granted is an Exchange permission the mailbox owner controls; a
    // 403 on their own mailbox is missing consent, which is an admin job.
    if (isShared) {
      return {
        content:
          `Access to ${mailbox.label} was denied. Knowva reads a shared mailbox as the user ` +
          "themselves, so this means this user hasn't been given access to that mailbox in " +
          "Exchange. Tell them that plainly, suggest they ask whoever administers the mailbox, " +
          "and offer to search their own mailbox instead.",
        isError: true,
      };
    }
    return {
      content:
        "Email search was denied (permission error). An administrator most likely needs to grant " +
        "consent for the Mail.Read permission, which Knowva needs to search mail. Tell the user " +
        "email search isn't set up yet -- and do not answer as though you had read their mail.",
      isError: true,
    };
  }

  if (status === 404) {
    return {
      content:
        isShared
          ? `${mailbox.label} does not exist, or is not a mailbox Knowva can reach. Tell the user ` +
            "you couldn't find that mailbox, and do not guess at its contents."
          : "That email could not be found. It may have been moved or deleted since it was " +
            "found in search. Tell the user that, and do not guess at what it said.",
      isError: true,
    };
  }

  if (status === 429) {
    return {
      content: "Email search is temporarily rate-limited. Tell the user to try again in a minute.",
      isError: true,
    };
  }

  return {
    content:
      `The email service returned an error${status ? ` (status ${status})` : ""}. ` +
      "Tell the user it's a problem on our side, not theirs, and do not guess at what any email " +
      "might have said.",
    isError: true,
  };
}

function statusOf(err: unknown): number | undefined {
  const e = err as { statusCode?: number; response?: { status?: number } };
  return typeof e?.statusCode === "number" ? e.statusCode : e?.response?.status;
}
