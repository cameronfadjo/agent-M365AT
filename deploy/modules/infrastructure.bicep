// ============================================================================
// Infrastructure Module — Azure Function App + Storage
// ============================================================================

param location string
param functionAppName string
param functionStorageAccountName string
param appServicePlanName string

@secure()
param azureOpenAIKey string
param azureOpenAIEndpoint string
param azureOpenAIDeployment string
param azureOpenAIDeploymentLarge string
param azureOpenAIApiVersion string

@secure()
param storageConnectionString string
param storageContainerName string
param sasTokenExpiryHours string

// ============================================================================
// STORAGE ACCOUNT (for Function App runtime — NOT the document storage)
// ============================================================================

resource functionStorage 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: functionStorageAccountName
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
  }
}

// ============================================================================
// APP SERVICE PLAN (Consumption / Serverless)
// ============================================================================

resource appServicePlan 'Microsoft.Web/serverfarms@2023-01-01' = {
  name: appServicePlanName
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  kind: 'linux'
  properties: {
    reserved: true  // Required for Linux
  }
}

// ============================================================================
// FUNCTION APP
// ============================================================================

resource functionApp 'Microsoft.Web/sites@2023-01-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  properties: {
    serverFarmId: appServicePlan.id
    reserved: true
    siteConfig: {
      pythonVersion: '3.11'
      linuxFxVersion: 'Python|3.11'
      appSettings: [
        // --- Function App Runtime ---
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${functionStorage.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${functionStorage.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTAZUREFILECONNECTIONSTRING'
          value: 'DefaultEndpointsProtocol=https;AccountName=${functionStorage.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${functionStorage.listKeys().keys[0].value}'
        }
        {
          name: 'WEBSITE_CONTENTSHARE'
          value: toLower(functionAppName)
        }
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'python'
        }
        {
          name: 'AzureWebJobsFeatureFlags'
          value: 'EnableWorkerIndexing'
        }

        // --- Azure OpenAI ---
        {
          name: 'AZURE_OPENAI_ENDPOINT'
          value: azureOpenAIEndpoint
        }
        {
          name: 'AZURE_OPENAI_KEY'
          value: azureOpenAIKey
        }
        {
          name: 'AZURE_OPENAI_DEPLOYMENT'
          value: azureOpenAIDeployment
        }
        {
          name: 'AZURE_OPENAI_DEPLOYMENT_LARGE'
          value: azureOpenAIDeploymentLarge
        }
        {
          name: 'AZURE_OPENAI_API_VERSION'
          value: azureOpenAIApiVersion
        }

        // --- Blob Storage (generated documents) ---
        {
          name: 'AZURE_STORAGE_CONNECTION_STRING'
          value: storageConnectionString
        }
        {
          name: 'AZURE_STORAGE_CONTAINER_NAME'
          value: storageContainerName
        }
        {
          name: 'SAS_TOKEN_EXPIRY_HOURS'
          value: sasTokenExpiryHours
        }

        // --- Entra ID (placeholders — set by deploy-entra.ps1) ---
        {
          name: 'ENTRA_TENANT_ID'
          value: 'PLACEHOLDER_SET_BY_ENTRA_SCRIPT'
        }
        {
          name: 'ENTRA_CLIENT_ID'
          value: 'PLACEHOLDER_SET_BY_ENTRA_SCRIPT'
        }
        {
          name: 'ENTRA_CLIENT_SECRET'
          value: 'PLACEHOLDER_SET_BY_ENTRA_SCRIPT'
        }
      ]
    }
    httpsOnly: true
  }
}

// ============================================================================
// OUTPUTS
// ============================================================================

output functionAppName string = functionApp.name
output functionAppDefaultHostName string = functionApp.properties.defaultHostName
output functionAppResourceId string = functionApp.id
