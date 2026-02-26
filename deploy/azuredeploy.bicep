// ============================================================================
// Refresh Agent — Azure Deployment Template (Bicep)
//
// Creates all Azure infrastructure for the Refresh Agent:
//   - Entra ID app registration (with OBO scopes + authorized clients)
//   - Azure Function App (Python 3.11, Linux, Consumption)
//   - Storage account (for Function App runtime)
//   - All environment variables / app settings
//
// Post-deployment manual steps:
//   1. Grant admin consent for Graph API permissions
//   2. Upload agent .zip in Teams Admin Center
//
// Author: Cameron Fadjo + Claude
// Version: 1.0
// ============================================================================

targetScope = 'subscription'

// ============================================================================
// PARAMETERS
// ============================================================================

@description('Azure region for all resources')
param location string

@description('Base name for all resources (e.g., "refresh-contoso"). Must be globally unique for the Function App.')
@minLength(3)
@maxLength(24)
param appName string

@description('Azure OpenAI endpoint URL (e.g., "https://myoai.openai.azure.com/")')
param azureOpenAIEndpoint string

@description('Azure OpenAI API key')
@secure()
param azureOpenAIKey string

@description('Azure OpenAI model deployment name')
param azureOpenAIDeployment string = 'gpt-4o-mini'

@description('Azure OpenAI large model deployment name (optional, for complex analysis)')
param azureOpenAIDeploymentLarge string = ''

@description('Azure OpenAI API version')
param azureOpenAIApiVersion string = '2025-01-01-preview'

@description('Azure Storage connection string for generated documents')
@secure()
param storageConnectionString string

@description('Blob container name for generated documents')
param storageContainerName string = 'generated-documents'

@description('SAS token expiry in hours')
param sasTokenExpiryHours string = '24'

// ============================================================================
// VARIABLES
// ============================================================================

var resourceGroupName = 'rg-${appName}'
var functionAppName = appName
var functionStorageAccountName = '${take(replace(toLower(appName), '-', ''), 20)}func'
var appServicePlanName = '${appName}-plan'
var functionAppDomain = '${functionAppName}.azurewebsites.net'

// Microsoft well-known client IDs for Teams/M365/Outlook
var authorizedClientAppIds = [
  '1fec8e78-bce4-4aaf-ab1b-5451cc387264'  // Teams Desktop & Mobile
  '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'  // Teams Web
  '4765445b-32c6-49b0-83e6-1115210e106b'  // Microsoft 365 (Office Web)
  '0ec893e0-5785-4de6-99da-4ed124e5296c'  // Microsoft 365 (Office Desktop)
  'd3590ed6-52b3-4102-aedd-aad2292ab01c'  // Outlook Desktop
  'bc59ab01-8403-45c6-8796-ac3ef710b3e3'  // Outlook Web
  '27922004-5251-4030-b22d-91ecd9a37ea4'  // Outlook Mobile
]

// ============================================================================
// RESOURCE GROUP
// ============================================================================

resource rg 'Microsoft.Resources/resourceGroups@2023-07-01' = {
  name: resourceGroupName
  location: location
}

// ============================================================================
// ENTRA ID APP REGISTRATION
// ============================================================================

// Note: Microsoft.Graph resources in Bicep require the Microsoft Graph Bicep
// extension (preview). If your environment doesn't support this, use the
// Deploy-RefreshAgent.ps1 script or the Azure CLI for Entra ID setup.
//
// The Entra ID configuration is handled by the companion deployment script
// (deploy-entra.ps1) which runs as a deployment script resource below.

// ============================================================================
// AZURE INFRASTRUCTURE (module)
// ============================================================================

module infrastructure 'modules/infrastructure.bicep' = {
  name: 'infrastructure'
  scope: rg
  params: {
    location: location
    functionAppName: functionAppName
    functionStorageAccountName: functionStorageAccountName
    appServicePlanName: appServicePlanName
    azureOpenAIEndpoint: azureOpenAIEndpoint
    azureOpenAIKey: azureOpenAIKey
    azureOpenAIDeployment: azureOpenAIDeployment
    azureOpenAIDeploymentLarge: azureOpenAIDeploymentLarge
    azureOpenAIApiVersion: azureOpenAIApiVersion
    storageConnectionString: storageConnectionString
    storageContainerName: storageContainerName
    sasTokenExpiryHours: sasTokenExpiryHours
  }
}

// ============================================================================
// OUTPUTS
// ============================================================================

output resourceGroupName string = rg.name
output functionAppName string = functionAppName
output functionAppUrl string = 'https://${functionAppDomain}'
output functionAppDomain string = functionAppDomain
output healthCheckUrl string = 'https://${functionAppDomain}/api/health'
output nextSteps string = 'Run deploy-entra.ps1 to configure Entra ID, then deploy Python code, then upload agent .zip'
