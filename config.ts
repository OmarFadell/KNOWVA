// Default Anthropic model. Override per environment with ANTHROPIC_MODEL so
// moving models never needs a code change.
const DEFAULT_ANTHROPIC_MODEL = "claude-sonnet-5";

const llmProvider = (process.env.LLM_PROVIDER || "anthropic").toLowerCase();

/**
 * Reads a required variable, throwing during module load so the process fails
 * to start. The alternative -- discovering the value is missing on the first
 * message -- surfaces as a confusing error for whoever happens to message the
 * bot first, potentially long after the deploy that caused it.
 */
function requireEnv(name: string, because: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(
      `Missing required environment variable ${name}: ${because}. ` +
        `Locally, set it in env/.env.local (or, for secrets, env/.env.local.user) and re-run ` +
        `Provision/Deploy so it is written into .localConfigs. ` +
        `On Azure, set ${name} as an App Service application setting.`
    );
  }
  return value;
}

/**
 * One entry from SHARED_MAILBOXES: a human-facing name the model can pass to
 * search_emails, and the SMTP address it resolves to.
 */
export interface SharedMailbox {
  /** Display name, e.g. "Project Alpha". Matched case-insensitively. */
  name: string;
  /** SMTP address the Graph call is made against, e.g. projectalpha@contoso.com. */
  address: string;
}

/**
 * Parses SHARED_MAILBOXES: a `;`-separated list of `Display Name=address` pairs.
 *
 *   SHARED_MAILBOXES="Project Alpha=projectalpha@contoso.com;Support=support@contoso.com"
 *
 * Optional -- unset means the email tools can only ever reach the asking user's
 * own mailbox, which is the intended default. But a *malformed* entry throws at
 * startup rather than being skipped: somebody who set this variable meant to
 * open a mailbox, and silently dropping their typo would leave them wondering
 * why the model keeps saying that mailbox does not exist.
 *
 * Naming a mailbox here grants nothing by itself. Every read still goes out on
 * the asking user's delegated token under Mail.Read.Shared, so Exchange decides
 * per user whether they may open it. This list only bounds *which* mailboxes
 * Knowva will ever ask for.
 */
function parseSharedMailboxes(raw: string | undefined): SharedMailbox[] {
  if (!raw || !raw.trim()) return [];

  return raw
    .split(";")
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const separator = entry.indexOf("=");
      const name = separator === -1 ? "" : entry.slice(0, separator).trim();
      const address = separator === -1 ? "" : entry.slice(separator + 1).trim();

      if (!name || !address || !address.includes("@")) {
        throw new Error(
          `Malformed SHARED_MAILBOXES entry "${entry}". Expected ` +
            `"Display Name=address@domain", with entries separated by ";" -- e.g. ` +
            `SHARED_MAILBOXES="Project Alpha=projectalpha@contoso.com;Support=support@contoso.com".`
        );
      }

      return { name, address };
    });
}

/**
 * Works out the public origin an OAuth callback will arrive on.
 *
 * ONE FUNCTION FOR EVERY PROVIDER, ON PURPOSE. This started as
 * resolveGitHubOrigin and was generalised when Atlassian became the second
 * OAuth provider, rather than copied. A second copy would have been a second
 * place for the redirect URI to drift away from the tunnel host -- which has
 * already cost this project once -- and a redirect URI that disagrees with the
 * registered one fails at the provider with a message that reads like a bad
 * client id.
 *
 * `overrideEnvName` is the per-provider escape hatch (GITHUB_OAUTH_ORIGIN,
 * ATLASSIAN_OAUTH_ORIGIN) for the case where a callback has to arrive somewhere
 * other than the bot's own endpoint. Everything else is shared: BOT_ENDPOINT
 * locally, BOT_DOMAIN on Azure.
 *
 * Returns "" when nothing is known, which each auth module reports as "sign-in
 * is not configured" rather than building a broken URL.
 */
function resolveOAuthOrigin(overrideEnvName: string): string {
  const explicit = process.env[overrideEnvName]?.trim();
  if (explicit) return explicit.replace(/\/+$/, "");

  // Local: the debug task writes the dev-tunnel URL here, scheme included.
  const endpoint = process.env.BOT_ENDPOINT?.trim();
  if (endpoint) return endpoint.replace(/\/+$/, "");

  // Azure: hostname only, no scheme.
  const domain = process.env.BOT_DOMAIN?.trim();
  if (domain) return `https://${domain.replace(/^https?:\/\//, "").replace(/\/+$/, "")}`;

  return "";
}

const config = {
  MicrosoftAppId: process.env.CLIENT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.TENANT_ID,
  MicrosoftAppPassword: process.env.CLIENT_PASSWORD,

  // The bot's own client secret, used to mint an *app-only* Graph token for
  // RSC chat reads (see src/graph/app-client.ts). Read from CLIENT_SECRET,
  // which is what the Toolkit actually writes into .localConfigs and what the
  // App Service setting in infra/azure.bicep is called. Not required to start:
  // a deployment without it simply cannot use search_conversations, and the
  // tool says so in plain language rather than the process refusing to boot.
  appClientSecret: process.env.CLIENT_SECRET || "",

  // Knowva's Teams app id, as written into appPackage/manifest.json's `id`.
  //
  // CURRENTLY UNUSED. It backed a /me/chats $filter that turned out not to be a
  // supported query (Graph 404s it) -- see src/graph/installed-chats.ts, which
  // now determines installation by attempting the message read instead. Kept
  // because the env plumbing is already in place and an installedApps-based
  // check would need it again.
  //
  // If it is ever used again: in Graph terms this value is a teamsApp
  // **externalId**, not a teamsApp **id**. `id` is the catalog-generated
  // identifier and is explicitly documented as different from the one in the
  // app package. Matching this against `teamsApp/id` returns nothing; match it
  // against `teamsApp/externalId`.
  teamsAppExternalId: process.env.TEAMS_APP_CATALOG_ID || process.env.TEAMS_APP_ID || "",

  llm: {
    provider: llmProvider,
    // Only the selected provider's credentials are demanded, so adding a
    // provider later does not force every environment to carry every key.
    anthropicApiKey:
      llmProvider === "anthropic"
        ? requireEnv("ANTHROPIC_API_KEY", "the Anthropic provider cannot authenticate without it")
        : "",
    anthropicModel: process.env.ANTHROPIC_MODEL || DEFAULT_ANTHROPIC_MODEL,
  },

  // The single SharePoint site the search_documents tool is scoped to, e.g.
  // https://contoso.sharepoint.com/sites/Knowledge. Required: grounded document
  // search is the whole point of this milestone, so a deployment without it is
  // misconfigured and should fail loudly at startup.
  sharePointSiteUrl: requireEnv(
    "SHAREPOINT_SITE_URL",
    "the document-search tool has no site to query without it"
  ),

  /**
   * GitHub, for the search_github tool. Wholly optional, and modelled on
   * appClientSecret above rather than on sharePointSiteUrl: a deployment with
   * no GitHub App configured starts fine and the tool says so in plain language,
   * because GitHub search is an addition rather than the reason Knowva exists.
   *
   * NOTE ON THE SECRET NAMES. The app reads GITHUB_CLIENT_ID /
   * GITHUB_CLIENT_SECRET, which m365agents.local.yml writes into .localConfigs
   * from SECRET_GITHUB_CLIENT_ID / SECRET_GITHUB_CLIENT_SECRET in
   * env/.env.local.user -- exactly the indirection ANTHROPIC_API_KEY already
   * uses. (A GitHub App client id is not actually a secret and could live in
   * env/.env.local as plain config; it is carried as a SECRET_ for now so both
   * halves of one credential sit in one place. Key Vault is the eventual home
   * for the secret half, same as the Anthropic key.)
   */
  github: {
    clientId: process.env.GITHUB_CLIENT_ID || "",
    clientSecret: process.env.GITHUB_CLIENT_SECRET || "",
    /**
     * Public origin the GitHub OAuth callback lands on, no trailing slash.
     *
     * Derived rather than configured, because it is already known twice over:
     * BOT_ENDPOINT is the dev-tunnel URL locally, BOT_DOMAIN is the App Service
     * hostname on Azure. GITHUB_OAUTH_ORIGIN overrides both, for the case where
     * the callback has to arrive somewhere other than the bot's own endpoint.
     *
     * THIS MUST MATCH A CALLBACK URL REGISTERED ON THE GITHUB APP, and the dev
     * tunnel hostname changes every time the tunnel is recreated -- see
     * src/auth/github.ts for what that means in practice.
     */
    oauthOrigin: resolveOAuthOrigin("GITHUB_OAUTH_ORIGIN"),
  },

  /**
   * Atlassian. PARKED -- read by src/auth/atlassian.ts, which app.ts does not
   * currently mount. The Xray milestone is org-scoped (one shared service
   * credential plus an admin-set project list), so there is no per-user Jira
   * sign-in to ask anybody for yet. Kept because the 3LO flow is complete and
   * verified, and a future "read Jira as the asking user" tool needs it.
   *
   * Optional in exactly the same way as `github` above: a deployment with no
   * Atlassian OAuth app starts fine.
   *
   * SECRET NAMES. The app reads ATLASSIAN_CLIENT_ID / ATLASSIAN_CLIENT_SECRET,
   * which m365agents.local.yml writes into .localConfigs from
   * SECRET_ATLASSIAN_CLIENT_ID / SECRET_ATLASSIAN_CLIENT_SECRET in
   * env/.env.local.user -- the same indirection GITHUB_CLIENT_ID already uses.
   */
  atlassian: {
    clientId: process.env.ATLASSIAN_CLIENT_ID || "",
    clientSecret: process.env.ATLASSIAN_CLIENT_SECRET || "",
    /**
     * Public origin the Atlassian OAuth callback lands on, no trailing slash.
     * Derived from BOT_ENDPOINT / BOT_DOMAIN by the shared resolver above --
     * NOT configured separately, and not derived independently.
     *
     * THIS MUST MATCH THE CALLBACK URL ON THE ATLASSIAN APP, and unlike a
     * GitHub App, an Atlassian OAuth 2.0 integration accepts only ONE. A
     * recreated dev tunnel therefore means editing that single field, not
     * adding to a list -- see src/auth/atlassian.ts.
     */
    oauthOrigin: resolveOAuthOrigin("ATLASSIAN_OAUTH_ORIGIN"),
  },

  /**
   * Xray Cloud, for the list_xray_projects tool.
   *
   * ONE SHARED SERVICE CREDENTIAL, AND THAT IS NOT AN INCONSISTENCY. Xray Cloud
   * authenticates with a Client ID / Client Secret pair from an API Key in its
   * Global Settings, and offers no per-user or delegated auth of any kind. So
   * unlike every other credential here, this one identifies the deployment
   * rather than a person.
   *
   * THE CREDENTIAL IS THE ENTIRE CONFIGURATION. There is no project list. An
   * earlier design took an admin-set XRAY_PROJECTS allowlist; it was removed
   * deliberately, and the tool now discovers projects from Xray itself. So what
   * Knowva reports is exactly what this API key can see -- no more, no less.
   * Whoever issues the key is making the disclosure decision, because Xray
   * applies that key's own permissions and nothing narrows it afterwards.
   *
   * XRAY_CLIENT_ID is plain config (env/.env.local, env/.env.dev); the secret
   * half comes from SECRET_XRAY_CLIENT_SECRET in env/.env.local.user locally
   * and from the xrayClientSecret @secure() Bicep parameter on Azure.
   *
   * THE BASE URL IS FIXED and not configurable. Xray Cloud does publish
   * regional hostnames (us./eu./au.xray.cloud.getxray.app), but the apex host
   * routes correctly for every tenant and one fixed value is one less thing to
   * get wrong per environment. It lives in src/xray/client.ts as a constant.
   */
  xray: {
    clientId: process.env.XRAY_CLIENT_ID || "",
    clientSecret: process.env.XRAY_CLIENT_SECRET || "",
  },

  // Optional, and deliberately so. With this unset, search_emails can only ever
  // read the asking user's own mailbox -- the default source, and the one that
  // needs no configuration. Each entry here adds an explicitly named shared or
  // project mailbox as an opt-in alternative scope the model may be pointed at.
  // Same shape as SHAREPOINT_SITE_URL: plain config, not a secret.
  sharedMailboxes: parseSharedMailboxes(process.env.SHARED_MAILBOXES),
};

export default config;
