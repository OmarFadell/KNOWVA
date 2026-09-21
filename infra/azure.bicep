@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@description('App Service plan SKU. F1 is free; switch to B1 for Always On.')
param webAppSku string = 'F1'

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param location string = resourceGroup().location

@maxLength(42)
param botDisplayName string

@description('Client ID of the Entra app created by aadApp/create (BOT_ID).')
param botAadAppClientId string

@secure()
@description('Client secret of that Entra app (SECRET_BOT_PASSWORD).')
param botAadAppClientSecret string

@description('Tenant ID. Switch to "common" when going multi-tenant.')
param botAadAppTenantId string

param oauthConnectionName string = 'graph'

@secure()
@description('Anthropic API key (SECRET_ANTHROPIC_API_KEY). Marked @secure() so it is redacted from deployment history and portal output.')
param anthropicApiKey string

@description('Claude model id. Overridable per environment so changing models needs no code change.')
param anthropicModel string = 'claude-sonnet-5'

@description('Teams app (catalog) id, from teamsApp/create. search_conversations needs it to work out which chats Knowva is installed in.')
param teamsAppId string

@minLength(1)
@description('Full URL of the SharePoint site the search_documents tool queries (SHAREPOINT_SITE_URL). Not a secret. Deployment fails if empty, because the app cannot start without it.')
param sharePointSiteUrl string

@description('GitHub App client id for the search_github tool (GITHUB_CLIENT_ID). Empty disables GitHub search; the tool then says so in plain language rather than the app failing to start.')
param githubClientId string = ''

@secure()
@description('GitHub App client secret (SECRET_GITHUB_CLIENT_SECRET). Marked @secure() so it is redacted from deployment history and portal output. Key Vault is the eventual home, same as the Anthropic key.')
param githubClientSecret string = ''

@description('Xray Cloud API key client id (XRAY_CLIENT_ID) for the list_xray_projects tool. ONE SHARED SERVICE CREDENTIAL for the whole deployment, not a per-user value -- Xray Cloud has no per-user auth model at all. Empty disables the tool, which then says so in plain language rather than the app failing to start.')
param xrayClientId string = ''

@secure()
@description('Xray Cloud API key client secret (SECRET_XRAY_CLIENT_SECRET). Marked @secure() so it is redacted from deployment history and portal output. Key Vault is the eventual home, same as the Anthropic key.')
param xrayClientSecret string = ''

@description('Optional shared/project mailboxes search_emails may be pointed at (SHARED_MAILBOXES), as "Display Name=address@domain" pairs separated by ";". Not a secret. Empty is the intended default: with nothing here, the email tools can only ever reach the mailbox of whoever is asking. Naming a mailbox grants nothing by itself -- every read still runs on that user delegated token under Mail.Read.Shared, so Exchange decides per user whether they may open it.')
param sharedMailboxes string = ''

// F1 (Free) does not support Always On -- setting it true makes the deployment fail.
var alwaysOnSupported = webAppSku != 'F1'

// The site's own public hostname, composed rather than read back from
// webApp.properties.defaultHostName, because a resource cannot reference its own
// properties from inside its own definition -- which is what an appSettings
// entry would be doing. The outputs at the bottom of this file use the real
// property; this is only for the setting the app reads at runtime.
//
// Assumes the public cloud's azurewebsites.net suffix, as does everything else
// here (see msaAppType: 'SingleTenant' in the bot registration). A sovereign
// cloud deployment would need this and the CLOUD env var changed together.
var webAppDomain = '${webAppName}.azurewebsites.net'

resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSku
  }
}

resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: alwaysOnSupported
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~22'
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'BOT_TYPE'
          value: 'SingleTenant'
        }
        {
          name: 'CLIENT_ID'
          value: botAadAppClientId
        }
        {
          name: 'CLIENT_SECRET'
          value: botAadAppClientSecret
        }
        {
          name: 'TENANT_ID'
          value: botAadAppTenantId
        }
        {
          name: 'AAD_APP_OAUTH_CONNECTION_NAME'
          value: oauthConnectionName
        }
        // The key lands in plain text in App Service application settings, which
        // anyone with read access to the site can see. The better long-term
        // answer is Key Vault: store the secret there and set this value to a
        // reference -- @Microsoft.KeyVault(SecretUri=...) -- with the web app
        // given a managed identity and a get-secret role assignment. Deferred
        // for now because the bot deliberately has no managed identity (an
        // MI-backed bot cannot carry the SSO OAuth connection).
        {
          name: 'ANTHROPIC_API_KEY'
          value: anthropicApiKey
        }
        {
          name: 'ANTHROPIC_MODEL'
          value: anthropicModel
        }
        {
          name: 'SHAREPOINT_SITE_URL'
          value: sharePointSiteUrl
        }
        {
          name: 'SHARED_MAILBOXES'
          value: sharedMailboxes
        }
        {
          name: 'GITHUB_CLIENT_ID'
          value: githubClientId
        }
        {
          name: 'GITHUB_CLIENT_SECRET'
          value: githubClientSecret
        }
        // Xray Cloud. One shared service credential, and it is the only bound
        // on what the tool reports -- see the parameter descriptions above for
        // why this one is deliberately not per-user.
        {
          name: 'XRAY_CLIENT_ID'
          value: xrayClientId
        }
        {
          name: 'XRAY_CLIENT_SECRET'
          value: xrayClientSecret
        }
        // The public origin the GitHub OAuth callback comes back to. config.ts
        // derives the full callback URL from it, and it must match a Callback
        // URL registered on the GitHub App.
        {
          name: 'BOT_DOMAIN'
          value: webAppDomain
        }
        {
          name: 'TEAMS_APP_ID'
          value: teamsAppId
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
}

module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    botDisplayName: botDisplayName
    botAadAppClientId: botAadAppClientId
    botAadAppClientSecret: botAadAppClientSecret
    botAadAppTenantId: botAadAppTenantId
    botAppDomain: webApp.properties.defaultHostName
    oauthConnectionName: oauthConnectionName
  }
}

// NOTE: BOT_ID is no longer an output here -- it now comes from aadApp/create
// in m365agents.yml. Two writers for one variable is how envs drift.
output AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
output BOT_SERVICE_NAME string = azureBotRegistration.outputs.BOT_SERVICE_NAME
output AAD_APP_OAUTH_CONNECTION_NAME string = azureBotRegistration.outputs.OAUTH_CONNECTION_NAME
