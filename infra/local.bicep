// Provisions the bot registration for the LOCAL environment.
// No App Service -- the bot points at the dev tunnel that the debug task starts.
// Because this redeploys on every provision, the messaging endpoint is always
// refreshed to the current tunnel host. No manual `az bot update` needed.

@maxLength(20)
@minLength(4)
param resourceBaseName string

@maxLength(42)
param botDisplayName string

param botAadAppClientId string

@secure()
param botAadAppClientSecret string

param botAadAppTenantId string

@description('Dev tunnel hostname, no scheme. Written to BOT_DOMAIN by the debug task.')
param botAppDomain string

param oauthConnectionName string = 'graph'

module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration-local'
  params: {
    resourceBaseName: resourceBaseName
    botDisplayName: botDisplayName
    botAadAppClientId: botAadAppClientId
    botAadAppClientSecret: botAadAppClientSecret
    botAadAppTenantId: botAadAppTenantId
    botAppDomain: botAppDomain
    oauthConnectionName: oauthConnectionName
  }
}

output BOT_SERVICE_NAME string = azureBotRegistration.outputs.BOT_SERVICE_NAME
output AAD_APP_OAUTH_CONNECTION_NAME string = azureBotRegistration.outputs.OAUTH_CONNECTION_NAME
