import {
  ClientSecretCredential,
  ManagedIdentityCredential,
  type AccessToken,
  type TokenCredential,
} from "@azure/identity";
import { Client as GraphClient } from "@microsoft/teams.graph";
import config from "../../config";

/**
 * An *application* Graph client -- the bot's own identity, with no user behind it.
 *
 * Every other Graph call in Knowva is delegated (the user's token from Teams
 * SSO, security-trimmed to what they can see). RSC is different by design:
 * ChatMessage.Read.Chat is an **Application** permission, granted per-chat at
 * install time, so reading a chat's messages requires an app-only token. There
 * is no delegated equivalent -- see the RSC permission table, where
 * ChatMessage.Read.Chat is "Application: Supported / Delegated: NA".
 *
 * This is NOT a tenant-wide key. Graph evaluates the RSC grant per chat id, so
 * `GET /chats/{id}/messages` with this token succeeds only for chats where
 * somebody installed Knowva, and 403s everywhere else. That property is what
 * makes "installation is the consent" enforceable rather than a promise: the
 * bot cannot read a chat it was never added to, no matter what it asks for.
 *
 * Because of that, the *enumeration* half deliberately does NOT use this client
 * -- see installed-chats.ts, which uses the asking user's delegated token so
 * results can never include a chat that user isn't a member of.
 */

const GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default";

// Refresh slightly before real expiry. A token that expires mid-request costs a
// user their answer; a few minutes of early refresh costs one extra token call.
const EXPIRY_SKEW_MS = 5 * 60 * 1000;

/**
 * Thrown when the process has no credentials to get an app token with. Callers
 * turn this into a "not configured" message rather than a generic Graph error,
 * because the fix is a deployment change, not a retry.
 */
export class AppCredentialsUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "AppCredentialsUnavailableError";
  }
}

let credential: TokenCredential | undefined;
let cachedToken: AccessToken | undefined;

function getCredential(): TokenCredential {
  if (credential) return credential;

  // Mirrors the two shapes app.ts already supports. UserAssignedMsi has no
  // secret; every other bot type authenticates with client id + secret.
  if (config.MicrosoftAppType === "UserAssignedMsi") {
    credential = new ManagedIdentityCredential({ clientId: config.MicrosoftAppId });
    return credential;
  }

  if (!config.MicrosoftAppId || !config.appClientSecret) {
    throw new AppCredentialsUnavailableError(
      "No application credentials available: CLIENT_ID and CLIENT_SECRET must both be set " +
        "to acquire an app-only Graph token for RSC chat reads."
    );
  }

  if (!config.MicrosoftAppTenantId) {
    throw new AppCredentialsUnavailableError(
      "No tenant id available: TENANT_ID must be set to acquire an app-only Graph token."
    );
  }

  credential = new ClientSecretCredential(
    config.MicrosoftAppTenantId,
    config.MicrosoftAppId,
    config.appClientSecret
  );
  return credential;
}

/** Cached app-only Graph token. Shared across conversations -- it is not user data. */
async function getAppToken(): Promise<string> {
  const now = Date.now();
  if (cachedToken && cachedToken.expiresOnTimestamp - EXPIRY_SKEW_MS > now) {
    return cachedToken.token;
  }

  const token = await getCredential().getToken(GRAPH_DEFAULT_SCOPE);
  if (!token) {
    throw new AppCredentialsUnavailableError(
      "The identity provider returned no app-only Graph token."
    );
  }

  cachedToken = token;
  return token.token;
}

/**
 * Builds a Graph client carrying the bot's app-only token.
 * Async because the token may need fetching; the client itself is cheap.
 */
export async function appGraphClient(): Promise<GraphClient> {
  return new GraphClient({ token: await getAppToken() });
}
