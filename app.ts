import { stripMentionsText, TokenCredentials, TypingActivityInput } from "@microsoft/teams.api";
import type { ActivityLike } from "@microsoft/teams.api";
import { App, ExpressAdapter } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import type { ILogger } from "@microsoft/teams.common";
import type { Client as GraphClient } from "@microsoft/teams.graph";
import config from "./config";
import { ManagedIdentityCredential } from "@azure/identity";
import { getMe } from "./src/graph/me";
import { graphClientForToken } from "./src/graph/client";
import { describeError } from "./src/errors";
import { createLlmProvider } from "./src/llm";
import type { LlmMessage } from "./src/llm";
import { appendToHistory, clearHistory, getHistory } from "./src/llm/history";
import { buildSystemPrompt } from "./src/llm/prompt";
import { runAgent } from "./src/agent";
import type { AgentToolSignal } from "./src/agent";
import { createSharePointSearchTool } from "./src/tools/sharepoint-search";
import { createDocumentMetadataTool } from "./src/tools/document-metadata";
import { createConversationSearchTool } from "./src/tools/conversation-search";
import { createEmailAccess } from "./src/tools/email-access";
import { createEmailSearchTool } from "./src/tools/email-search";
import { createEmailContentTool } from "./src/tools/email-content";
import { createGitHubSearchTool } from "./src/tools/github";
import { createGitHubContentTool } from "./src/tools/github-content";
import { createActingIdentityResolver } from "./src/auth/acting-identity";
import {
  buildSignInUrl,
  connectedLogin,
  isGitHubConfigured,
  registerGitHubOAuthRoutes,
  signOut,
} from "./src/auth/github";
import { gitHubSignInCard } from "./src/teams/github-signin-card";
import {
  disableEmailForUser,
  enableEmailForUser,
  isEmailDisabledForUser,
} from "./src/user/email-preferences";
import {
  groupChatWelcomeMessage,
  personalWelcomeMessage,
  helpMessage,
  isGroupConversation,
  pauseConversation,
  resumeConversation,
  shouldIngest,
  shouldRespond,
} from "./src/teams/group-chat";
import {
  clearAmbientMessages,
  formatAmbientContext,
  recordAmbientMessage,
} from "./src/teams/chat-context";
import { withSendRetry } from "./src/send-retry";

// Built once at startup. config.ts has already validated the credentials for
// the selected provider by this point, so a misconfigured deployment fails here
// rather than on somebody's first message.
const llm = createLlmProvider();

// Create storage for conversation history
const storage = new LocalStorage();

const createTokenFactory = () => {
  return async (scope: string | string[], tenantId?: string): Promise<string> => {
    const managedIdentityCredential = new ManagedIdentityCredential({
      clientId: process.env.CLIENT_ID,
    });
    const scopes = Array.isArray(scope) ? scope : [scope];
    const tokenResponse = await managedIdentityCredential.getToken(scopes, {
      tenantId: tenantId,
    });

    return tokenResponse.token;
  };
};

// Configure authentication using TokenCredentials
const tokenCredentials: TokenCredentials = {
  clientId: process.env.CLIENT_ID || "",
  token: createTokenFactory(),
};

const credentialOptions =
  config.MicrosoftAppType === "UserAssignedMsi" ? { ...tokenCredentials } : undefined;

// The HTTP layer, held explicitly rather than left to the App to construct.
//
// This is the same ExpressAdapter the App would have created for itself -- same
// HttpServer, same /api/messages registration, same Bot Framework JWT
// validation. The only difference is that we keep the reference, which is what
// makes it possible to mount the GitHub OAuth callback on the same server and
// the same port. Nothing about Teams SSO changes.
const httpAdapter = new ExpressAdapter();

// Create the app with storage
const app = new App({
  ...credentialOptions,
  httpServerAdapter: httpAdapter,
  storage,
  skipAuth: !process.env.CLIENT_ID,
  oauth: {
    // Read from env rather than hardcoding "graph" -- must match the OAuth
    // connection name on the Azure Bot resource (see infra/botRegistration/azurebot.bicep).
    defaultConnectionName: process.env.AAD_APP_OAUTH_CONNECTION_NAME || "graph",
  },
});

// The GitHub OAuth callback, mounted on the same server and port as
// /api/messages but NOT behind the Teams JWT middleware -- the caller is a
// browser following a redirect from github.com and has no Teams token to
// present. src/auth/github.ts explains why that is safe (single-use,
// server-side, short-lived state) and no-ops when no GitHub App is configured.
registerGitHubOAuthRoutes(httpAdapter, app.log);

// Interface for conversation state
interface ConversationState {
  count: number;
}

const getConversationState = (conversationId: string): ConversationState => {
  let state = storage.get(conversationId);
  if (!state) {
    state = { count: 0 };
    storage.set(conversationId, state);
  }
  return state;
};

// A message that arrived before the user had a token. The SSO token exchange
// completes on a *separate* invoke activity, so the text has to survive the gap
// between "exchange started" and the signin event that finishes the turn.
interface PendingSignin {
  text: string;
  count: number;
  /**
   * The ambient group-chat context as it stood when the question was asked.
   * Captured rather than recomputed, so the answer is grounded in what the chat
   * looked like at the time -- a busy group chat can move on considerably in the
   * seconds a token exchange takes.
   */
  conversationContext?: string | null;
  tenantId?: string;
  /**
   * The asker's Entra object id, captured on the message turn. Carried across
   * the exchange because the email opt-out is checked against it, and the
   * signin event's activity is a different activity from the one that asked the
   * question -- reading `from` off the wrong one would check the wrong person.
   */
  aadObjectId?: string;
  /** Whether the question was asked in a group chat, where replies are public. */
  inGroupChat?: boolean;
}

// Keyed per user, not per conversation: in a group chat several people can each
// have an exchange in flight at the same time.
const pendingKey = (conversationId: string, userId: string) =>
  `pending-signin:${conversationId}:${userId}`;

/**
 * Identifies the user via Graph, then answers with the LLM.
 * Takes the pieces it needs rather than a whole context, so it can serve both
 * the message turn and the signin event, whose contexts differ.
 */
async function respondWithLlm(
  graph: GraphClient,
  log: ILogger,
  send: (activity: ActivityLike) => Promise<unknown>,
  text: string,
  conversationId: string,
  options: {
    /** Attributed recent messages from this group chat, or null in a 1:1. */
    conversationContext?: string | null;
    /** Used to build message deep links that resolve for multi-tenant users. */
    tenantId?: string;
    /**
     * `activity.from.aadObjectId` for whoever asked. The email tools check the
     * per-user opt-out against this as well as against the identity they read
     * back from the token itself -- see src/tools/email-access.ts.
     */
    aadObjectId?: string;
    /** Replies are visible to the whole room, which changes the email rules. */
    inGroupChat?: boolean;
  } = {}
): Promise<void> {
  // Teams shows nothing at all while we wait on Graph and the model, and a model
  // turn runs to several seconds. This has to go out before any slow work.
  // Best-effort: a typing indicator that fails to send must never cost the user
  // their answer, which is what an unhandled throw here would do.
  try {
    await send(new TypingActivityInput());
  } catch (err) {
    log.warn("send: typing indicator failed, continuing", describeError(err));
  }

  // The SSO payoff: the display name goes into the system prompt so Claude can
  // address the user directly. Losing the name should degrade the reply, not
  // cost the user their answer, so this failure is non-fatal.
  let displayName = "the user";
  let actingUserId: string | undefined;
  try {
    const me = await getMe(graph);
    displayName = me.displayName;
    // Reused as the acting identity for the email opt-out check, purely to save
    // the email tools a second /me round trip. If this call failed, they resolve
    // it themselves rather than skipping the check -- losing the name may degrade
    // a reply, but losing the identity must not weaken a privacy control.
    actingUserId = me.id;
  } catch (err) {
    log.error("graph: /me failed, answering without the user's name", describeError(err), err);
  }

  const userMessage: LlmMessage = { role: "user", content: text };

  // Both tools are delegated and security-trim to what this user can see, so
  // they are built around *this turn's* Graph client (the user's token), not a
  // shared one. Claude chains them: search_documents finds a document and
  // returns its webUrl, get_document_metadata resolves that URL when the
  // question is about currency or authorship.
  //
  // search_conversations is the odd one out: it takes the user's client to work
  // out *which* chats to look in, but reads the messages with the bot's own
  // app-only token, because RSC grants are application permissions. See
  // src/tools/conversation-search.ts for why it needs both.
  //
  // The email tools are delegated too, and emphatically only delegated: there is
  // no app-only mail path anywhere, because an application Mail.Read grant is
  // tenant-wide over every mailbox in the org. So search_emails can reach this
  // user's mailbox and nothing else -- which is also why the opt-out below can
  // check "the person chatting" and be checking the right person.
  //
  // The two email tools share one EmailAccess: it memoises the acting identity
  // and the opt-out decision for the turn, so a search-then-read chain costs one
  // identity resolution rather than two, and cannot answer the "may we?"
  // question differently on the second call than it did on the first.
  const emailAccess = createEmailAccess({
    graph,
    chattingUserId: options.aadObjectId,
    actingUserId,
    log,
  });

  // search_github needs the same question answered -- whose credentials is this
  // turn allowed to spend? -- but the stakes differ. For email, Graph refuses
  // anyway if we get it wrong. For GitHub the token IS the authority, so this
  // resolver cross-checks the token's own identity against the one Teams put on
  // the activity and refuses if they disagree. See src/auth/acting-identity.ts.
  // Both GitHub tools share it, so a search-then-read chain resolves once.
  const gitHubIdentity = createActingIdentityResolver({
    graph,
    claimedUserId: options.aadObjectId,
    resolvedUserId: actingUserId,
    log,
  });

  const tools = [
    createSharePointSearchTool(graph, config.sharePointSiteUrl, log),
    createDocumentMetadataTool(graph, log),
    createConversationSearchTool(graph, options.tenantId, log),
    createEmailSearchTool(graph, emailAccess, log),
    createEmailContentTool(graph, emailAccess, log),
    createGitHubSearchTool(gitHubIdentity, log),
    createGitHubContentTool(gitHubIdentity, log),
  ];

  try {
    const result = await runAgent({
      provider: llm,
      system: buildSystemPrompt(displayName, {
        conversationContext: options.conversationContext,
        inGroupChat: options.inGroupChat,
      }),
      // History is committed only after a successful reply, so a failed turn
      // cannot leave an unanswered user message behind to confuse the next one.
      // The loop's intermediate tool turns stay local to runAgent and are never
      // persisted.
      messages: [...getHistory(conversationId), userMessage],
      tools,
      logger: log,
      maxIterations: 5,
    });

    if (result.stopReason === "refusal") {
      log.warn(`llm: ${llm.name}/${llm.model} declined this request`);
      await send("I can't help with that one, sorry. Try rephrasing it, or ask me something else.");
      return;
    }

    if (result.stopReason === "max_tokens") {
      log.warn(`llm: ${llm.name}/${llm.model} hit the output cap; reply may be cut short`);
    }

    const reply = result.text || "I couldn't put a reply together for that. Try asking again.";
    appendToHistory(conversationId, userMessage, { role: "assistant", content: reply });
    await send(reply);

    // THE ONE PLACE A TOOL GETS TO DRIVE THE UI.
    //
    // A tool that needs the user to authorize something cannot say so in prose
    // alone: the authorize URL is single-use, expires in minutes, and belongs
    // to one person, so it must not pass through the model. search_github
    // therefore raises a signal (see AgentToolSignal in src/agent/loop.ts) and
    // the card is attached here, after the model's own explanation.
    //
    // Sent as a second activity rather than merged into the reply because the
    // reply has already gone out by this point -- and because a card that
    // arrives under the explanation reads as an answer to it.
    await sendAuthCards(result.signals, send, options.aadObjectId, log);
  } catch (err) {
    log.error(`llm: ${llm.name}/${llm.model} request failed`, describeError(err), err);
    await send(
      "I couldn't reach my language model just then -- that's a problem on my side, not yours. " +
        "Give it a moment and try again."
    );
  }
}

/**
 * Turns `needs-auth` signals into sign-in cards.
 *
 * Best-effort throughout: the user already has the model's explanation, so a
 * card that cannot be built or cannot be sent must degrade to "no button"
 * rather than throwing away the answer that was already delivered.
 */
async function sendAuthCards(
  signals: AgentToolSignal[],
  send: (activity: ActivityLike) => Promise<unknown>,
  aadObjectId: string | undefined,
  log: ILogger
): Promise<void> {
  for (const signal of signals) {
    if (signal.kind !== "needs-auth" || signal.provider !== "github") continue;

    // The sign-in URL is bound to a specific user, so without an identity there
    // is nobody to bind it to. The tool has already told the model to explain
    // the situation; this just declines to offer a button that could not work.
    if (!aadObjectId) {
      log.warn("github-auth: cannot offer a sign-in card without an aadObjectId for the user");
      continue;
    }

    const url = buildSignInUrl(aadObjectId);
    if (!url) {
      log.warn("github-auth: sign-in requested but no GitHub App is configured");
      continue;
    }

    try {
      await send(gitHubSignInCard(url));
    } catch (err) {
      log.error("github-auth: failed to send the sign-in card", describeError(err), err);
    }
  }
}

app.on("message", async (context) => {
  const activity = context.activity;
  const text: string = stripMentionsText(activity);
  const conversationId = activity.conversation.id;
  const inGroup = isGroupConversation(activity);

  // READING AND ANSWERING ARE TWO DIFFERENT DECISIONS.
  //
  // In a group chat Knowva receives every message, mentioned or not -- that is
  // what the groupChat scope buys, and it is what makes a later "so what did we
  // land on?" answerable. It is NOT permission to talk. So this handler ingests
  // first, unconditionally, and only then asks whether it was addressed.
  //
  // Note the ordering: context is captured BEFORE this message is recorded, so
  // the model does not get the question it is answering handed back to it twice.
  const conversationContext = inGroup ? formatAmbientContext(conversationId) : null;

  if (shouldIngest(activity, text)) {
    recordAmbientMessage(conversationId, {
      authorName: activity.from.name?.trim() || "an unidentified participant",
      // tagOnly keeps the *inner* text of every mention, so "@Priya can you
      // check this" is remembered as "Priya can you check this" rather than
      // losing the name entirely. Who was being addressed is often the whole
      // point of a message. The bot's own question text above still uses the
      // full strip, because Knowva does not need its own name echoed back.
      text: stripMentionsText(activity, { tagOnly: true }),
      timestamp: (activity.timestamp ? new Date(activity.timestamp) : new Date()).toISOString(),
    });
  }

  // The soft off switch. Deliberately handled before the shouldRespond gate:
  // /resume has to work while paused, or the pause would be permanent.
  if (text === "/pause") {
    pauseConversation(conversationId);
    await context.send(
      "Ok -- I've stopped reading this conversation. Say **/resume** to turn me back on. " +
        "If you want me gone for good, remove me from the chat: that revokes my access rather " +
        "than just asking me not to look."
    );
    return;
  }

  if (text === "/resume") {
    resumeConversation(conversationId);
    await context.send("Ok -- I'm reading this conversation again.");
    return;
  }

  // THE EMAIL OPT-OUT, handled here with /pause rather than below with /reset.
  //
  // Two reasons for the position. It has to work while a conversation is
  // paused, exactly as /resume does -- a privacy control you can lock yourself
  // out of is not one. And it must not need an @mention in a group chat: this
  // is about somebody's mailbox, and making them address the bot in front of
  // the room to switch it off is the wrong shape entirely.
  //
  // The flag is per user and global, not per conversation: it is keyed on the
  // Entra object id, so running this anywhere disables email search for that
  // person everywhere. See src/user/email-preferences.ts.
  if (text === "/disable-email" || text === "/enable-email") {
    const userId = activity.from.aadObjectId;

    if (!userId) {
      // Rare -- Teams omits aadObjectId on some surfaces. Refusing is the only
      // honest answer: writing the flag under a Teams-surface id would record a
      // preference the Graph-side check could never find, which is worse than
      // saying it did not work.
      context.log.warn(
        "email-preferences: no aadObjectId on the activity; cannot record the preference " +
          `(conversation=${conversationId} user=${activity.from.id})`
      );
      await context.send(
        "I couldn't work out who you are just then, so I haven't changed anything. " +
          "Try again in a moment -- and if it keeps happening, message me directly rather than " +
          "in a group chat."
      );
      return;
    }

    if (text === "/disable-email") {
      disableEmailForUser(userId);
      context.log.info(`email-preferences: email search disabled for user ${userId}`);
      await context.send(
        "Done -- I won't search your email any more. That applies everywhere, not just in this " +
          "chat, and it takes effect before I make any request to Outlook. Say **/enable-email** " +
          "to turn it back on.\n\n" +
          "One honest caveat: I keep this setting in memory, so it can be forgotten if I'm " +
          "restarted or redeployed. If that matters to you, check with **/disable-email** again " +
          "after an update."
      );
      return;
    }

    enableEmailForUser(userId);
    context.log.info(`email-preferences: email search re-enabled for user ${userId}`);
    await context.send(
      "Ok -- I can search your email again. That covers your Inbox and Sent Items only; " +
        "never drafts, deleted mail, or anyone else's mailbox. **/disable-email** switches it " +
        "back off."
    );
    return;
  }

  // Everything past here is a reply. In a group chat that means Knowva was
  // @mentioned; other people's messages stop here, having been remembered.
  if (!shouldRespond(activity)) {
    return;
  }

  // GitHub sign-out. Sits with the email opt-out rather than with /reset for the
  // same reasons: it is a credential control, so it must work in a paused
  // conversation and must not require an @mention in front of a room.
  //
  // NOTE THE LIMIT, which the reply states plainly: this forgets Knowva's copy
  // of the token, it does not revoke the authorization at GitHub. Only the user
  // can do that, from their GitHub settings, and pretending otherwise would be
  // the kind of half-truth that matters for a credential.
  if (text === "/github-signout") {
    const userId = activity.from.aadObjectId;

    if (!userId) {
      await context.send(
        "I couldn't work out who you are just then, so I haven't changed anything. Try again in " +
          "a moment."
      );
      return;
    }

    const had = signOut(userId);
    context.log.info(
      `github-auth: sign-out requested by ${userId} (had a session: ${had})`
    );

    await context.send(
      (had
        ? "Done -- I've forgotten your GitHub connection and won't search GitHub as you any more."
        : "You weren't connected to GitHub, so there was nothing to disconnect.") +
        "\n\nTo fully revoke my access, remove the KnowvaGithubApp authorization in your " +
        "GitHub settings -- signing out here only clears my copy of the token."
    );
    return;
  }

  if (text === "/help") {
    const userId = activity.from.aadObjectId;
    await context.send(
      helpMessage(isEmailDisabledForUser(userId), {
        configured: isGitHubConfigured(),
        login: userId ? connectedLogin(userId) : undefined,
      })
    );
    return;
  }

  if (text === "/reset") {
    storage.delete(conversationId);
    clearHistory(conversationId);
    // Also drop the ambient buffer: "forget what we discussed" plainly covers
    // the surrounding chat Knowva was holding, not just its own dialogue.
    clearAmbientMessages(conversationId);
    await context.send("Ok I've cleared this conversation -- I've forgotten what we discussed.");
    return;
  }

  if (text === "/count") {
    const state = getConversationState(activity.conversation.id);
    await context.send(`The count is ${state.count}`);
    return;
  }

  if (text === "/diag") {
    await context.send(JSON.stringify(activity));
    return;
  }

  if (text === "/state") {
    const state = getConversationState(activity.conversation.id);
    await context.send(JSON.stringify(state));
    return;
  }

  if (text === "/runtime") {
    const runtime = {
      nodeversion: process.version,
      sdkversion: "2.0.0", // Microsoft Teams SDK
    };
    await context.send(JSON.stringify(runtime));
    return;
  }

  // Default behaviour: answer with the LLM, identified via silent Teams SSO.
  const state = getConversationState(conversationId);
  state.count++;

  // Needed to build message deep links that land in the right tenant for a
  // guest or multi-tenant user. Falls back to the bot's own configured tenant.
  const tenantId = activity.conversation.tenantId || config.MicrosoftAppTenantId;

  // Fast path: the eager token fetch at the start of the turn found a cached
  // token, so the exchange already happened on some earlier turn.
  if (context.isSignedIn && context.userToken) {
    await respondWithLlm(
      context.userGraph,
      context.log,
      withSendRetry((a) => context.send(a), context.log),
      text,
      conversationId,
      {
        conversationContext,
        tenantId,
        aadObjectId: activity.from.aadObjectId,
        inGroupChat: inGroup,
      }
    );
    return;
  }

  // No cached token. Teams SSO is not passive -- it never hands over a token
  // unprompted. The bot has to send an OAuth card carrying a tokenExchangeResource,
  // which is what ctx.signin() does; Teams then performs the exchange silently and
  // posts back a signin/tokenExchange invoke. The SDK answers that invoke itself and
  // emits the "signin" event handled below. Without this call nothing ever asks
  // Teams for a token, so the token store stays empty and every turn falls back.
  const key = pendingKey(conversationId, activity.from.id);
  storage.set(key, {
    text,
    count: state.count,
    conversationContext,
    tenantId,
    aadObjectId: activity.from.aadObjectId,
    inGroupChat: inGroup,
  } as PendingSignin);

  try {
    const token = await context.signin();

    if (token) {
      // signin() found a token in the store after all. context.userGraph closed
      // over this turn's (empty) eager fetch, so it can't see this token -- build
      // a client around the token we were just handed.
      storage.delete(key);
      await respondWithLlm(
        graphClientForToken(token),
        context.log,
        withSendRetry((a) => context.send(a), context.log),
        text,
        conversationId,
        {
          conversationContext,
          tenantId,
          aadObjectId: activity.from.aadObjectId,
          inGroupChat: inGroup,
        }
      );
      return;
    }

    // Card sent, exchange in flight. The signin event finishes this turn.
    // Distinct from a failure: nothing has gone wrong yet.
    context.log.info(
      `sso: token exchange initiated on connection "${context.connectionName}" ` +
        `(conversation=${activity.conversation.id} user=${activity.from.id} ` +
        `conversationType=${activity.conversation.conversationType ?? "personal"})`
    );
  } catch (err) {
    storage.delete(key);
    context.log.error(
      `sso: could not initiate token exchange on connection "${context.connectionName}" ` +
        `(conversation=${activity.conversation.id} user=${activity.from.id})`,
      describeError(err),
      err
    );
    await context.send(
      `[${state.count}] you said: ${text}\n\n` +
        "(I couldn't start sign-in just now. That's a problem on my side, not yours -- it's been logged.)"
    );
  }
});

// Fires when Knowva is added to a conversation. In a group chat this is the
// consent moment: the install itself granted the ChatMessage.Read.Chat RSC
// permission declared in appPackage/manifest.json, Teams has already posted its
// own "<person> added Knowva" system message naming who did it, and Knowva is
// now in the member roster where everyone can see it.
//
// Nobody clicks anything else, which is exactly why this message exists. It is
// a disclosure, not a request: by the time it sends, access has already been
// granted, so its only job is to make sure nobody in the chat is surprised
// later about what Knowva can see -- and to name uninstalling as the way out.
app.on("install.add", async (context) => {
  const activity = context.activity;

  // Personal installs used to be skipped here, on the grounds that the 1:1
  // experience was covered by the app's own store description. Outlook mail
  // changed that: the description does not mention email, nobody clicks
  // anything to grant it, and the opt-out is worthless if the only people who
  // learn about it are the ones who already knew to ask. So both surfaces now
  // get a disclosure, and both name /disable-email and /help.
  const inGroup = isGroupConversation(activity);

  context.log.info(
    `install: added to a ${inGroup ? "group" : "personal"} conversation ` +
      `(conversation=${activity.conversation.id} by=${activity.from.id})`
  );

  try {
    await withSendRetry((a: ActivityLike) => context.send(a), context.log)(
      inGroup ? groupChatWelcomeMessage() : personalWelcomeMessage()
    );
  } catch (err) {
    // Worth an error, not a crash. The install succeeded and RSC is granted
    // either way -- but the chat now has a bot reading it that never announced
    // itself, which is precisely the situation the message exists to prevent.
    context.log.error(
      `install: failed to post the welcome disclosure to ${activity.conversation.id}`,
      describeError(err),
      err
    );
  }
});

// Fires once Teams has completed the SSO token exchange (or an interactive
// sign-in). The SDK's built-in signin/tokenExchange route performs the exchange
// and hands over an already-authenticated userGraph, so all that's left is to
// finish the turn the message handler had to suspend.
app.event("signin", async (context) => {
  const activity = context.activity;
  const key = pendingKey(activity.conversation.id, activity.from.id);
  const pending: PendingSignin | undefined = storage.get(key);
  storage.delete(key);

  context.log.info(
    `sso: token exchange completed on connection "${context.connectionName}" ` +
      `(conversation=${activity.conversation.id} user=${activity.from.id})`
  );

  if (!pending) {
    // Sign-in completed with no message waiting on it -- e.g. a duplicate
    // exchange, or a turn that already resolved. Nothing to echo.
    context.log.debug("sso: no pending message for completed sign-in");
    return;
  }

  await respondWithLlm(
    context.userGraph,
    context.log,
    withSendRetry((a) => context.send(a), context.log),
    pending.text,
    activity.conversation.id,
    {
      conversationContext: pending.conversationContext,
      tenantId: pending.tenantId,
      // From the message turn, not from this signin activity. They are different
      // activities and only the first one identifies who actually asked.
      aadObjectId: pending.aadObjectId,
      inGroupChat: pending.inGroupChat,
    }
  );
});

// Teams reports token-exchange problems through signin/failure, which the SDK
// turns into an error event (codes like resourcematchfailed = OAuth card token
// exchange URL does not match the Application ID URI, or tokenmissing). Logging
// the whole object here is the difference between a diagnosable failure and a
// silent one.
app.event("error", ({ error, activity }) => {
  app.log.error(
    `app error${activity ? ` while handling ${activity.type}` : ""}`,
    describeError(error),
    error
  );
});

export default app;
