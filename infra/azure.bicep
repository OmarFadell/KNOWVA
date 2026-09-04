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

// F1 (Free) does not support Always On -- setting it true makes the deployment fail.
var alwaysOnSupported = webAppSku != 'F1'

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
