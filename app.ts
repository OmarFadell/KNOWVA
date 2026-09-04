import { stripMentionsText, TokenCredentials } from "@microsoft/teams.api";
import { App } from "@microsoft/teams.apps";
import { LocalStorage } from "@microsoft/teams.common";
import type { ILogger } from "@microsoft/teams.common";
import type { Client as GraphClient } from "@microsoft/teams.graph";
import config from "./config";
import { ManagedIdentityCredential } from "@azure/identity";
import { getMe } from "./src/graph/me";
import { graphClientForToken } from "./src/graph/client";
import { describeError } from "./src/errors";

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
}

// Keyed per user, not per conversation: in a group chat several people can each
// have an exchange in flight at the same time.
const pendingKey = (conversationId: string, userId: string) =>
  `pending-signin:${conversationId}:${userId}`;

/**
 * Looks the user up in Graph and sends the identity-aware echo.
 * Takes the pieces it needs rather than a whole context, so it can serve both
 * the message turn and the signin event, whose contexts differ.
 */
async function replyWithIdentity(
  graph: GraphClient,
  log: ILogger,
  send: (text: string) => Promise<unknown>,
  text: string,
  count: number
): Promise<void> {
  try {
    const me = await getMe(graph);
    await send(`Hi ${me.displayName}! [${count}] you said: ${text}`);
  } catch (err) {
    log.error("graph: /me failed after a successful token exchange", describeError(err), err);
    await send(
      `[${count}] you said: ${text}\n\n` +
        "(I know who you are but couldn't reach Microsoft Graph to look up your name -- try again shortly.)"
    );
  }
}

app.on("message", async (context) => {
  const activity = context.activity;
  const text: string = stripMentionsText(activity);

  if (text === "/reset") {
    storage.delete(activity.conversation.id);
    await context.send("Ok I've deleted the current conversation state.");
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

  // Default echo behavior, identity-aware via silent Teams SSO.
  const state = getConversationState(activity.conversation.id);
  state.count++;

  // Fast path: the eager token fetch at the start of the turn found a cached
  // token, so the exchange already happened on some earlier turn.
  if (context.isSignedIn && context.userToken) {
    await replyWithIdentity(
      context.userGraph,
      context.log,
      (t) => context.send(t),
      text,
      state.count
    );
    return;
  }

  // No cached token. Teams SSO is not passive -- it never hands over a token
  // unprompted. The bot has to send an OAuth card carrying a tokenExchangeResource,
  // which is what ctx.signin() does; Teams then performs the exchange silently and
  // posts back a signin/tokenExchange invoke. The SDK answers that invoke itself and
  // emits the "signin" event handled below. Without this call nothing ever asks
  // Teams for a token, so the token store stays empty and every turn falls back.
  const key = pendingKey(activity.conversation.id, activity.from.id);
  storage.set(key, { text, count: state.count } as PendingSignin);

  try {
    const token = await context.signin();

    if (token) {
      // signin() found a token in the store after all. context.userGraph closed
      // over this turn's (empty) eager fetch, so it can't see this token -- build
      // a client around the token we were just handed.
      storage.delete(key);
      await replyWithIdentity(
        graphClientForToken(token),
        context.log,
        (t) => context.send(t),
        text,
        state.count
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

  await replyWithIdentity(
    context.userGraph,
    context.log,
    (t) => context.send(t),
    pending.text,
    pending.count
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
