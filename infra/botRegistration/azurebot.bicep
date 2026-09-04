@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@maxLength(42)
param botDisplayName string

param botServiceName string = resourceBaseName
param botServiceSku string = 'F0'

@description('Client ID of the Entra app registration backing the bot (BOT_ID).')
param botAadAppClientId string

@secure()
@description('Client secret of that Entra app (SECRET_BOT_PASSWORD). Needed by the OAuth connection.')
param botAadAppClientSecret string

@description('Tenant ID the bot is scoped to. Switch to "common" when going multi-tenant.')
param botAadAppTenantId string

@description('Hostname only, no scheme. e.g. knowva.azurewebsites.net or abc-3978.uks1.devtunnels.ms')
param botAppDomain string

@description('Name of the OAuth connection used for Teams SSO. Must match AAD_APP_OAUTH_CONNECTION_NAME in code.')
param oauthConnectionName string = 'graph'

@description('Graph scopes granted through the OAuth connection. Milestone 1 only needs User.Read.')
param oauthScopes string = 'User.Read'

// SingleTenant (not UserAssignedMSI) -- an MI-backed bot cannot carry an OAuth
// connection setting, which Teams SSO token exchange requires.
resource botService 'Microsoft.BotService/botServices@2022-09-15' = {
  kind: 'azurebot'
  location: 'global'
  name: botServiceName
  sku: {
    name: botServiceSku
  }
  properties: {
    displayName: botDisplayName
    endpoint: 'https://${botAppDomain}/api/messages'
    msaAppId: botAadAppClientId
    msaAppTenantId: botAadAppTenantId
    msaAppType: 'SingleTenant'
  }
}

resource botServiceMsTeamsChannel 'Microsoft.BotService/botServices/channels@2022-09-15' = {
  parent: botService
  location: 'global'
  name: 'MsTeamsChannel'
  properties: {
    channelName: 'MsTeamsChannel'
  }
}

// The OAuth connection Teams SSO exchanges its token against.
// serviceProviderId is the well-known ID for the "Azure Active Directory v2" provider.
// Verify for your cloud with:
//   az bot authsetting list-providers --query "[?properties.displayName=='Azure Active Directory v2'].id"
resource graphConnection 'Microsoft.BotService/botServices/connections@2022-09-15' = {
  parent: botService
  location: 'global'
  name: oauthConnectionName
  properties: {
    serviceProviderId: '30dd229c-58e3-4a48-bdfd-91ec48eb906c'
    serviceProviderDisplayName: 'Azure Active Directory v2'
    clientId: botAadAppClientId
    clientSecret: botAadAppClientSecret
    scopes: oauthScopes
    parameters: [
      {
        key: 'tenantID'
        value: botAadAppTenantId
      }
      {
        key: 'tokenExchangeUrl'
        value: 'api://botid-${botAadAppClientId}'
      }
    ]
  }
}

output BOT_DOMAIN string = botAppDomain
output BOT_SERVICE_NAME string = botService.name
output OAUTH_CONNECTION_NAME string = graphConnection.name
