<#
.SYNOPSIS
    Automated deployment script for the Refresh Agent (M365 Agents Toolkit Edition).
    Configures Entra ID, Azure Function App, environment variables, and deploys code.

.DESCRIPTION
    This script automates Parts 2-4 of the Deployment Guide:
      - Part 2: Entra ID app registration (permissions, scopes, authorized clients, secret)
      - Part 3: Azure Function App creation, environment config, code deployment
      - Part 4: Placeholder replacement in agent manifest files

    Part 5 (Agents Toolkit provisioning) remains manual because it requires
    the VS Code extension and interactive M365 login.

.PARAMETER TenantId
    The Azure AD / Entra ID tenant ID (GUID or domain like contoso.onmicrosoft.com)

.PARAMETER SubscriptionId
    Azure subscription ID for creating the Function App and storage resources

.PARAMETER AppName
    Base name for all resources (e.g., "refresh-contoso"). Used to derive:
      - Function App name:   {AppName}
      - Resource group:      rg-{AppName}
      - Entra app name:      Refresh Agent - {AppName}

.PARAMETER Region
    Azure region for the Function App (e.g., "eastus", "westus2", "centralus")

.PARAMETER AzureOpenAIEndpoint
    Full URL to your Azure OpenAI resource (e.g., "https://myoai.openai.azure.com/")

.PARAMETER AzureOpenAIKey
    API key for your Azure OpenAI resource

.PARAMETER AzureOpenAIDeployment
    Deployment name for the GPT model (default: "gpt-4o-mini")

.PARAMETER AzureOpenAIApiVersion
    Azure OpenAI API version (default: "2025-01-01-preview")

.PARAMETER StorageConnectionString
    Azure Storage connection string for blob storage (generated documents)

.PARAMETER StorageContainerName
    Blob container name (default: "generated-documents")

.PARAMETER ProjectPath
    Path to the Refresh-M365AT project folder (default: current directory)

.PARAMETER SkipFunctionDeploy
    Skip Azure Function deployment (useful if rerunning just the Entra/config steps)

.PARAMETER SkipEntraSetup
    Skip Entra ID setup (useful if app registration already exists)

.EXAMPLE
    .\Deploy-RefreshAgent.ps1 `
        -TenantId "contoso.onmicrosoft.com" `
        -SubscriptionId "12345678-abcd-efgh-ijkl-123456789012" `
        -AppName "refresh-contoso" `
        -Region "eastus" `
        -AzureOpenAIEndpoint "https://contoso-oai.openai.azure.com/" `
        -AzureOpenAIKey "sk-abc123..." `
        -StorageConnectionString "DefaultEndpointsProtocol=https;AccountName=..."

.NOTES
    Prerequisites:
      - Azure CLI (az) installed and in PATH
      - Azure Functions Core Tools v4 (func) installed and in PATH
      - Microsoft Graph PowerShell SDK: Install-Module Microsoft.Graph -Scope CurrentUser
      - PowerShell 7+ recommended

    Author: Cameron Fadjo + Claude
    Version: 1.0
    Date: February 2026
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$SubscriptionId,

    [Parameter(Mandatory = $true)]
    [string]$AppName,

    [Parameter(Mandatory = $true)]
    [string]$Region,

    [Parameter(Mandatory = $true)]
    [string]$AzureOpenAIEndpoint,

    [Parameter(Mandatory = $true)]
    [string]$AzureOpenAIKey,

    [Parameter(Mandatory = $false)]
    [string]$AzureOpenAIDeployment = "gpt-4o-mini",

    [Parameter(Mandatory = $false)]
    [string]$AzureOpenAIDeploymentLarge = "",

    [Parameter(Mandatory = $false)]
    [string]$AzureOpenAIApiVersion = "2025-01-01-preview",

    [Parameter(Mandatory = $true)]
    [string]$StorageConnectionString,

    [Parameter(Mandatory = $false)]
    [string]$StorageContainerName = "generated-documents",

    [Parameter(Mandatory = $false)]
    [string]$SasTokenExpiryHours = "24",

    [Parameter(Mandatory = $false)]
    [string]$ProjectPath = ".",

    [switch]$SkipFunctionDeploy,

    [switch]$SkipEntraSetup
)

# ============================================================================
# CONFIGURATION
# ============================================================================

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$ResourceGroupName = "rg-$AppName"
$FunctionAppName = $AppName
$FunctionAppDomain = "$FunctionAppName.azurewebsites.net"
$FunctionAppUrl = "https://$FunctionAppDomain"
$EntraAppName = "Refresh Agent - $AppName"
$AzureFunctionPath = Join-Path $ProjectPath "azure_function"
$AppPackagePath = Join-Path $ProjectPath "appPackage"
$OpenApiPath = Join-Path $AppPackagePath "apiSpecificationFile" "openapi.yaml"
$ManifestPath = Join-Path $AppPackagePath "manifest.json"

# Microsoft well-known client IDs for authorized client applications
$AuthorizedClientApps = @(
    @{ Id = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"; Name = "Teams Desktop & Mobile" },
    @{ Id = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"; Name = "Teams Web" },
    @{ Id = "4765445b-32c6-49b0-83e6-1115210e106b"; Name = "Microsoft 365 (Office Web)" },
    @{ Id = "0ec893e0-5785-4de6-99da-4ed124e5296c"; Name = "Microsoft 365 (Office Desktop)" },
    @{ Id = "d3590ed6-52b3-4102-aedd-aad2292ab01c"; Name = "Outlook Desktop" },
    @{ Id = "bc59ab01-8403-45c6-8796-ac3ef710b3e3"; Name = "Outlook Web" },
    @{ Id = "27922004-5251-4030-b22d-91ecd9a37ea4"; Name = "Outlook Mobile" }
)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-Step {
    param([string]$Step, [string]$Message)
    Write-Information ""
    Write-Information "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Write-Information "  [$Step] $Message"
    Write-Information "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
}

function Write-SubStep {
    param([string]$Message)
    Write-Information "    → $Message"
}

function Write-Success {
    param([string]$Message)
    Write-Information "    ✓ $Message"
}

function Write-Warn {
    param([string]$Message)
    Write-Warning "    ⚠ $Message"
}

function Test-CommandExists {
    param([string]$Command)
    $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
}

# ============================================================================
# PREREQUISITES CHECK
# ============================================================================

Write-Step "0/6" "Checking prerequisites"

# Check Azure CLI
if (-not (Test-CommandExists "az")) {
    throw "Azure CLI (az) is not installed. Install from https://aka.ms/installazurecli"
}
Write-Success "Azure CLI found"

# Check Azure Functions Core Tools
if (-not (Test-CommandExists "func")) {
    throw "Azure Functions Core Tools (func) not found. Install: npm install -g azure-functions-core-tools@4"
}
Write-Success "Azure Functions Core Tools found"

# Check project structure
if (-not (Test-Path $AzureFunctionPath)) {
    throw "azure_function/ directory not found at $AzureFunctionPath. Ensure ProjectPath points to the Refresh-M365AT folder."
}
if (-not (Test-Path (Join-Path $AzureFunctionPath "host.json"))) {
    throw "host.json not found in azure_function/. Project structure appears incomplete."
}
if (-not (Test-Path $OpenApiPath)) {
    throw "openapi.yaml not found at $OpenApiPath. Ensure appPackage/apiSpecificationFile/openapi.yaml exists."
}
Write-Success "Project structure verified"

# Check Microsoft.Graph module
$graphModule = Get-Module -ListAvailable -Name "Microsoft.Graph.Applications" -ErrorAction SilentlyContinue
if (-not $graphModule -and -not $SkipEntraSetup) {
    Write-Warn "Microsoft.Graph PowerShell module not found."
    Write-Information "    Installing Microsoft.Graph.Applications module..."
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force -AllowClobber
    Write-Success "Microsoft.Graph.Applications module installed"
}
elseif ($graphModule) {
    Write-Success "Microsoft.Graph PowerShell module found"
}

# ============================================================================
# STEP 1: AZURE LOGIN
# ============================================================================

Write-Step "1/6" "Authenticating to Azure"

Write-SubStep "Logging into Azure CLI..."
az login --tenant $TenantId 2>&1 | Out-Null
az account set --subscription $SubscriptionId 2>&1 | Out-Null

$currentAccount = az account show --query "{name:name, id:id}" -o json | ConvertFrom-Json
Write-Success "Logged in: $($currentAccount.name) ($($currentAccount.id))"

# ============================================================================
# STEP 2: ENTRA ID APP REGISTRATION (Deployment Guide Part 2)
# ============================================================================

$EntraClientId = $null
$EntraClientSecret = $null

if (-not $SkipEntraSetup) {
    Write-Step "2/6" "Setting up Entra ID app registration"

    # Connect to Microsoft Graph
    Write-SubStep "Connecting to Microsoft Graph..."
    Connect-MgGraph -TenantId $TenantId -Scopes "Application.ReadWrite.All" -NoWelcome

    # --- 2.1 Create App Registration ---
    Write-SubStep "Creating app registration: $EntraAppName"

    $existingApp = Get-MgApplication -Filter "displayName eq '$EntraAppName'" -ErrorAction SilentlyContinue
    if ($existingApp) {
        Write-Warn "App '$EntraAppName' already exists (ID: $($existingApp.AppId)). Using existing."
        $app = $existingApp
    }
    else {
        $appParams = @{
            DisplayName    = $EntraAppName
            SignInAudience = "AzureADMyOrg"
        }
        $app = New-MgApplication @appParams
        Write-Success "App registered: $($app.AppId)"
    }

    $EntraClientId = $app.AppId
    $EntraObjectId = $app.Id
    $TenantGuid = (Get-MgContext).TenantId

    # --- 2.2 Configure API Permissions ---
    Write-SubStep "Configuring API permissions (Files.ReadWrite, User.Read)..."

    # Microsoft Graph App ID
    $graphResourceId = "00000003-0000-0000-c000-000000000000"

    # Permission IDs for delegated permissions
    $filesReadWrite = "5c28f0bf-8a70-41f1-8ee2-e11b8e0ee7c7"  # Files.ReadWrite
    $userRead = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"         # User.Read

    $requiredAccess = @{
        ResourceAppId  = $graphResourceId
        ResourceAccess = @(
            @{ Id = $filesReadWrite; Type = "Scope" },
            @{ Id = $userRead; Type = "Scope" }
        )
    }

    Update-MgApplication -ApplicationId $EntraObjectId -RequiredResourceAccess @($requiredAccess)
    Write-Success "API permissions configured"

    # Grant admin consent via Azure CLI (Graph PowerShell doesn't easily do this)
    Write-SubStep "Granting admin consent..."
    # Create service principal first if it doesn't exist
    $sp = Get-MgServicePrincipal -Filter "appId eq '$EntraClientId'" -ErrorAction SilentlyContinue
    if (-not $sp) {
        $sp = New-MgServicePrincipal -AppId $EntraClientId
    }
    # Admin consent via az cli
    az ad app permission admin-consent --id $EntraClientId 2>&1 | Out-Null
    Write-Success "Admin consent granted"

    # --- 2.3 Expose an API ---
    Write-SubStep "Configuring Application ID URI and access_as_user scope..."

    $appIdUri = "api://$FunctionAppDomain/$EntraClientId"

    # Define the OAuth2 permission scope
    $scope = @{
        AdminConsentDescription = "Allows M365 Copilot to call the Refresh API on behalf of the signed-in user"
        AdminConsentDisplayName = "Access Refresh API as user"
        Id                      = [guid]::NewGuid().ToString()
        IsEnabled               = $true
        Type                    = "User"
        UserConsentDescription  = "Allow Refresh Agent to access your OneDrive files"
        UserConsentDisplayName  = "Access Refresh API"
        Value                   = "access_as_user"
    }

    $apiSettings = @{
        IdentifierUris       = @($appIdUri)
        Api = @{
            Oauth2PermissionScopes = @($scope)
        }
    }

    Update-MgApplication -ApplicationId $EntraObjectId @apiSettings
    Write-Success "Application ID URI set: $appIdUri"
    Write-Success "access_as_user scope created"

    # --- 2.4 Authorize Client Applications ---
    Write-SubStep "Adding authorized client applications..."

    $scopeId = $scope.Id
    $preAuthorizedApps = @()

    foreach ($clientApp in $AuthorizedClientApps) {
        $preAuthorizedApps += @{
            AppId                  = $clientApp.Id
            DelegatedPermissionIds = @($scopeId)
        }
        Write-Success "Authorized: $($clientApp.Name) ($($clientApp.Id))"
    }

    $apiUpdate = @{
        Api = @{
            Oauth2PermissionScopes = @($scope)
            PreAuthorizedApplications = $preAuthorizedApps
        }
    }

    Update-MgApplication -ApplicationId $EntraObjectId @apiUpdate
    Write-Success "All client applications authorized"

    # --- 2.5 Create Client Secret ---
    Write-SubStep "Creating client secret..."

    $secretParams = @{
        PasswordCredential = @{
            DisplayName = "Refresh Agent Auto-Deploy"
            EndDateTime = (Get-Date).AddMonths(24)
        }
    }

    $secret = Add-MgApplicationPassword -ApplicationId $EntraObjectId @secretParams
    $EntraClientSecret = $secret.SecretText
    Write-Success "Client secret created (expires: $($secret.EndDateTime))"

    # --- 2.6 Summary ---
    Write-Information ""
    Write-Information "    ┌─────────────────────────────────────────────────┐"
    Write-Information "    │  Entra ID Configuration Complete                │"
    Write-Information "    ├─────────────────────────────────────────────────┤"
    Write-Information "    │  Client ID:     $EntraClientId"
    Write-Information "    │  Tenant ID:     $TenantGuid"
    Write-Information "    │  App ID URI:    $appIdUri"
    Write-Information "    │  Secret:        ****$(($EntraClientSecret).Substring($EntraClientSecret.Length - 4))"
    Write-Information "    └─────────────────────────────────────────────────┘"
}
else {
    Write-Step "2/6" "Skipping Entra ID setup (--SkipEntraSetup)"
    # Prompt for values if skipping
    if (-not $EntraClientId) {
        $EntraClientId = Read-Host "Enter ENTRA_CLIENT_ID"
    }
    if (-not $EntraClientSecret) {
        $EntraClientSecret = Read-Host "Enter ENTRA_CLIENT_SECRET" -AsSecureString
        $EntraClientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($EntraClientSecret)
        )
    }
    $TenantGuid = $TenantId
}

# ============================================================================
# STEP 3: AZURE FUNCTION APP (Deployment Guide Part 3)
# ============================================================================

Write-Step "3/6" "Setting up Azure Function App"

# --- 3.1 Create Resource Group ---
Write-SubStep "Creating resource group: $ResourceGroupName in $Region"
az group create --name $ResourceGroupName --location $Region --output none 2>&1
Write-Success "Resource group ready"

# --- 3.1 Create Function App ---
Write-SubStep "Creating Function App: $FunctionAppName"

$existingFunc = az functionapp show --name $FunctionAppName --resource-group $ResourceGroupName --query "name" -o tsv 2>$null
if ($existingFunc) {
    Write-Warn "Function App '$FunctionAppName' already exists. Using existing."
}
else {
    # Create storage account for the Function App runtime (separate from doc storage)
    $storageAccountName = ($AppName -replace '[^a-z0-9]', '').Substring(0, [Math]::Min(20, ($AppName -replace '[^a-z0-9]', '').Length)) + "func"
    Write-SubStep "Creating storage account: $storageAccountName"
    az storage account create `
        --name $storageAccountName `
        --resource-group $ResourceGroupName `
        --location $Region `
        --sku Standard_LRS `
        --output none 2>&1

    Write-SubStep "Creating Function App..."
    az functionapp create `
        --name $FunctionAppName `
        --resource-group $ResourceGroupName `
        --storage-account $storageAccountName `
        --consumption-plan-location $Region `
        --runtime python `
        --runtime-version 3.11 `
        --os-type Linux `
        --functions-version 4 `
        --output none 2>&1
}
Write-Success "Function App ready: $FunctionAppUrl"

# --- 3.2 Configure Environment Variables ---
Write-SubStep "Configuring environment variables (12 settings)..."

$appSettings = @(
    "AzureWebJobsFeatureFlags=EnableWorkerIndexing",
    "AZURE_OPENAI_ENDPOINT=$AzureOpenAIEndpoint",
    "AZURE_OPENAI_KEY=$AzureOpenAIKey",
    "AZURE_OPENAI_DEPLOYMENT=$AzureOpenAIDeployment",
    "AZURE_OPENAI_API_VERSION=$AzureOpenAIApiVersion",
    "AZURE_STORAGE_CONNECTION_STRING=$StorageConnectionString",
    "AZURE_STORAGE_CONTAINER_NAME=$StorageContainerName",
    "SAS_TOKEN_EXPIRY_HOURS=$SasTokenExpiryHours",
    "ENTRA_TENANT_ID=$TenantGuid",
    "ENTRA_CLIENT_ID=$EntraClientId",
    "ENTRA_CLIENT_SECRET=$EntraClientSecret"
)

# Add large model deployment if specified
if ($AzureOpenAIDeploymentLarge) {
    $appSettings += "AZURE_OPENAI_DEPLOYMENT_LARGE=$AzureOpenAIDeploymentLarge"
}

az functionapp config appsettings set `
    --name $FunctionAppName `
    --resource-group $ResourceGroupName `
    --settings @appSettings `
    --output none 2>&1

Write-Success "All environment variables configured"

# ============================================================================
# STEP 4: DEPLOY AZURE FUNCTIONS (Deployment Guide Part 3.4)
# ============================================================================

if (-not $SkipFunctionDeploy) {
    Write-Step "4/6" "Deploying Azure Functions"

    Write-SubStep "Publishing from: $AzureFunctionPath"

    Push-Location $AzureFunctionPath
    try {
        $deployOutput = func azure functionapp publish $FunctionAppName 2>&1
        $deployOutput | ForEach-Object { Write-Information "    $_" }
    }
    finally {
        Pop-Location
    }

    # Verify deployment
    Write-SubStep "Verifying deployment (waiting 15 seconds for cold start)..."
    Start-Sleep -Seconds 15

    try {
        $healthResponse = Invoke-RestMethod -Uri "$FunctionAppUrl/api/health" -Method Get -TimeoutSec 30
        if ($healthResponse.status -eq "healthy") {
            $endpointCount = ($healthResponse.endpoints | Measure-Object).Count
            Write-Success "Health check passed: $endpointCount endpoints registered"
        }
        else {
            Write-Warn "Health check returned unexpected status: $($healthResponse.status)"
        }
    }
    catch {
        Write-Warn "Health check failed: $($_.Exception.Message)"
        Write-Warn "Functions may still be starting up. Check manually: $FunctionAppUrl/api/health"
    }
}
else {
    Write-Step "4/6" "Skipping Function deployment (--SkipFunctionDeploy)"
}

# ============================================================================
# STEP 5: STAMP PLACEHOLDERS (Deployment Guide Part 4.1)
# ============================================================================

Write-Step "5/6" "Replacing placeholders in agent manifests"

# --- openapi.yaml ---
Write-SubStep "Updating openapi.yaml..."
$openapiContent = Get-Content $OpenApiPath -Raw
$openapiContent = $openapiContent -replace '\$\{\{AZURE_FUNCTION_URL\}\}', $FunctionAppUrl
$openapiContent = $openapiContent -replace '\$\{\{ENTRA_TENANT_ID\}\}', $TenantGuid
Set-Content -Path $OpenApiPath -Value $openapiContent -NoNewline
Write-Success "openapi.yaml: AZURE_FUNCTION_URL and ENTRA_TENANT_ID replaced"

# --- manifest.json ---
Write-SubStep "Updating manifest.json..."
$manifestContent = Get-Content $ManifestPath -Raw
$manifestContent = $manifestContent -replace '\$\{\{AZURE_FUNCTION_DOMAIN\}\}', $FunctionAppDomain
$manifestContent = $manifestContent -replace '\$\{\{ENTRA_CLIENT_ID\}\}', $EntraClientId
Set-Content -Path $ManifestPath -Value $manifestContent -NoNewline
Write-Success "manifest.json: AZURE_FUNCTION_DOMAIN and ENTRA_CLIENT_ID replaced"

# Note: TEAMS_APP_ID, OAUTH2_REGISTRATION_ID, APP_NAME_SUFFIX left as ${{...}}
# for the Agents Toolkit to populate during provisioning
Write-Success "Auto-populated placeholders left intact for Agents Toolkit"

# ============================================================================
# STEP 6: SUMMARY
# ============================================================================

Write-Step "6/6" "Deployment complete"

Write-Information ""
Write-Information "  ┌─────────────────────────────────────────────────────────────┐"
Write-Information "  │                  REFRESH AGENT DEPLOYED                      │"
Write-Information "  ├─────────────────────────────────────────────────────────────┤"
Write-Information "  │                                                              │"
Write-Information "  │  Function App:     $FunctionAppUrl"
Write-Information "  │  Health Check:     $FunctionAppUrl/api/health"
Write-Information "  │  Resource Group:   $ResourceGroupName"
Write-Information "  │  Entra Client ID:  $EntraClientId"
Write-Information "  │  Tenant ID:        $TenantGuid"
Write-Information "  │                                                              │"
Write-Information "  ├─────────────────────────────────────────────────────────────┤"
Write-Information "  │  NEXT STEPS (Manual — requires VS Code):                    │"
Write-Information "  │                                                              │"
Write-Information "  │  1. Open the project in VS Code:                            │"
Write-Information "  │     code $ProjectPath                                       │"
Write-Information "  │                                                              │"
Write-Information "  │  2. Click Agents Toolkit → Lifecycle → Provision            │"
Write-Information "  │     - Select 'dev' environment                              │"
Write-Information "  │     - Sign in to your M365 tenant                           │"
Write-Information "  │     - When prompted for OAuth client ID,                    │"
Write-Information "  │       enter: $EntraClientId"
Write-Information "  │                                                              │"
Write-Information "  │  3. Click Run and Debug → Preview in Copilot (Edge)         │"
Write-Information "  │                                                              │"
Write-Information "  │  4. Test: 'I need this year's back-to-school letter'        │"
Write-Information "  │                                                              │"
Write-Information "  └─────────────────────────────────────────────────────────────┘"

# ============================================================================
# EXPORT VALUES (for reuse or debugging)
# ============================================================================

$deploymentResult = [PSCustomObject]@{
    TenantId             = $TenantGuid
    SubscriptionId       = $SubscriptionId
    ResourceGroup        = $ResourceGroupName
    FunctionAppName      = $FunctionAppName
    FunctionAppUrl       = $FunctionAppUrl
    FunctionAppDomain    = $FunctionAppDomain
    EntraClientId        = $EntraClientId
    EntraAppName         = $EntraAppName
    AppIdUri             = "api://$FunctionAppDomain/$EntraClientId"
    Region               = $Region
    DeployedAt           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
}

# Save deployment values to a JSON file for reference
$outputPath = Join-Path $ProjectPath "deployment-output.json"
$deploymentResult | ConvertTo-Json -Depth 3 | Set-Content -Path $outputPath
Write-Information ""
Write-Information "  Deployment values saved to: $outputPath"
Write-Information "  (Keep this file — contains values needed for troubleshooting)"
Write-Information ""

# Return the result object
return $deploymentResult
