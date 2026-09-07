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

const config = {
  MicrosoftAppId: process.env.CLIENT_ID,
  MicrosoftAppType: process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.TENANT_ID,
  MicrosoftAppPassword: process.env.CLIENT_PASSWORD,

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
};

export default config;
