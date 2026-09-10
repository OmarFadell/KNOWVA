import { stripMentionsText, TokenCredentials, TypingActivityInput } from "@microsoft/teams.api";
import type { ActivityLike } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
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
import { createSharePointSearchTool } from "./src/tools/sharepoint-search";
import { createDocumentMetadataTool } from "./src/tools/document-metadata";
import { createConversationSearchTool } from "./src/tools/conversation-search";
import {
  groupChatWelcomeMessage,
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

// Create the app with storage
const app = new App({
  ...credentialOptions,
  storage,
  skipAuth: !process.env.CLIENT_ID,
  oauth: {
    // Read from env rather than hardcoding "graph" -- must match the OAuth
    // connection name on the Azure Bot resource (see infra/botRegistration/azurebot.bicep).
    defaultConnectionName: process.env.AAD_APP_OAUTH_CONNECTION_NAME || "graph",
  },
});

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
  try {
    const me = await getMe(graph);
    displayName = me.displayName;
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
  const tools = [
    createSharePointSearchTool(graph, config.sharePointSiteUrl, log),
    createDocumentMetadataTool(graph, log),
    createConversationSearchTool(graph, options.tenantId, log),
  ];

  try {
    const result = await runAgent({
      provider: llm,
      system: buildSystemPrompt(displayName, {
        conversationContext: options.conversationContext,
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
  } catch (err) {
    log.error(`llm: ${llm.name}/${llm.model} request failed`, describeError(err), err);
    await send(
      "I couldn't reach my language model just then -- that's a problem on my side, not yours. " +
        "Give it a moment and try again."
    );
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

  // Everything past here is a reply. In a group chat that means Knowva was
  // @mentioned; other people's messages stop here, having been remembered.
  if (!shouldRespond(activity)) {
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
      { conversationContext, tenantId }
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
        { conversationContext, tenantId }
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

  // Personal installs are the existing 1:1 experience and already covered by
  // the app's own description; only the group case is a disclosure to a room of
  // people who did not individually opt in.
  if (!isGroupConversation(activity)) return;

  context.log.info(
    `install: added to a group conversation ` +
      `(conversation=${activity.conversation.id} by=${activity.from.id})`
  );

  try {
    await withSendRetry((a: ActivityLike) => context.send(a), context.log)(
      groupChatWelcomeMessage()
    );
  } catch (err) {
    // Worth an error, not a crash. The install succeeded and RSC is granted
    // either way -- but the chat now has a bot reading it that never announced
    // itself, which is precisely the situation the message exists to prevent.
    context.log.error(
      `install: failed to post the group-chat disclosure to ${activity.conversation.id}`,
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
    { conversationContext: pending.conversationContext, tenantId: pending.tenantId }
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
