<#
.SYNOPSIS
    Full end-to-end deployment of the Refresh Agent into a customer tenant.
    Runs all steps: Azure infrastructure, Entra ID, code deployment, agent package.

.DESCRIPTION
    This is the master orchestrator. It runs:
      1. ARM/Bicep deployment (Azure Function App + storage)
      2. Entra ID configuration (app registration, permissions, scopes, secret)
      3. Python code deployment to Azure Function App
      4. Agent package build (.zip ready for sideloading)

    After this script completes, the customer admin needs to:
      - Upload the .zip in Teams Admin Center (or sideload in Teams)
      - That's it.

.EXAMPLE
    .\Deploy-Full.ps1 `
        -AppName "refresh-contoso" `
        -Region "eastus" `
        -AzureOpenAIEndpoint "https://contoso-oai.openai.azure.com/" `
        -AzureOpenAIKey "sk-..." `
        -StorageConnectionString "DefaultEndpointsProtocol=https;..."
#>

[CmdletBinding()]
param(
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
    [string]$SasTokenExpiryHours = "24"
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path $ScriptDir -Parent
$ResourceGroupName = "rg-$AppName"
$FunctionAppDomain = "$AppName.azurewebsites.net"

Write-Information ""
Write-Information "  ╔══════════════════════════════════════════════════════════╗"
Write-Information "  ║          REFRESH AGENT — FULL DEPLOYMENT                ║"
Write-Information "  ║          Target: $AppName"
Write-Information "  ╚══════════════════════════════════════════════════════════╝"
Write-Information ""

# ============================================================================
# PREREQUISITES
# ============================================================================

Write-Information "  [0/5] Checking prerequisites..."

foreach ($cmd in @("az", "func")) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        throw "'$cmd' is not installed. See DEPLOYMENT_GUIDE.md Prerequisites."
    }
}

$graphMod = Get-Module -ListAvailable -Name "Microsoft.Graph.Applications" -ErrorAction SilentlyContinue
if (-not $graphMod) {
    Write-Information "    Installing Microsoft.Graph.Applications..."
    Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force -AllowClobber
}
Write-Information "    ✓ All prerequisites met"

# ============================================================================
# STEP 1: AZURE LOGIN
# ============================================================================

Write-Information ""
Write-Information "  [1/5] Authenticating to Azure..."
Write-Information "    A browser window will open for Azure login."
az login 2>&1 | Out-Null

$account = az account show -o json | ConvertFrom-Json
$subscriptionId = $account.id
Write-Information "    ✓ Subscription: $($account.name) ($subscriptionId)"

# ============================================================================
# STEP 2: DEPLOY AZURE INFRASTRUCTURE
# ============================================================================

Write-Information ""
Write-Information "  [2/5] Deploying Azure infrastructure (Bicep)..."

$bicepPath = Join-Path $ScriptDir "azuredeploy.bicep"

az deployment sub create `
    --location $Region `
    --template-file $bicepPath `
    --parameters `
        location=$Region `
        appName=$AppName `
        azureOpenAIEndpoint=$AzureOpenAIEndpoint `
        azureOpenAIKey=$AzureOpenAIKey `
        azureOpenAIDeployment=$AzureOpenAIDeployment `
        azureOpenAIDeploymentLarge=$AzureOpenAIDeploymentLarge `
        azureOpenAIApiVersion=$AzureOpenAIApiVersion `
        storageConnectionString=$StorageConnectionString `
        storageContainerName=$StorageContainerName `
        sasTokenExpiryHours=$SasTokenExpiryHours `
    --output none 2>&1

Write-Information "    ✓ Resource group, Function App, and storage created"

# ============================================================================
# STEP 3: CONFIGURE ENTRA ID
# ============================================================================

Write-Information ""
Write-Information "  [3/5] Configuring Entra ID..."

$entraResult = & (Join-Path $ScriptDir "deploy-entra.ps1") -AppName $AppName -ResourceGroupName $ResourceGroupName

$ClientId = $entraResult.ClientId
$TenantId = $entraResult.TenantId

Write-Information "    ✓ Entra ID configured"

# ============================================================================
# STEP 4: DEPLOY PYTHON CODE
# ============================================================================

Write-Information ""
Write-Information "  [4/5] Deploying Python code to Azure Function App..."

$azureFunctionPath = Join-Path $ProjectRoot "azure_function"
Push-Location $azureFunctionPath
try {
    func azure functionapp publish $AppName 2>&1 | ForEach-Object {
        if ($_ -match "error|fail" -and $_ -notmatch "ErrorAction") {
            Write-Warning "    $_"
        }
    }
}
finally {
    Pop-Location
}

# Verify
Write-Information "    Waiting 15 seconds for cold start..."
Start-Sleep -Seconds 15

try {
    $health = Invoke-RestMethod -Uri "https://$FunctionAppDomain/api/health" -TimeoutSec 30
    $endpointCount = ($health.endpoints | Measure-Object).Count
    Write-Information "    ✓ Health check passed: $endpointCount endpoints"
}
catch {
    Write-Warning "    ⚠ Health check inconclusive — functions may still be starting"
    Write-Warning "    Check manually: https://$FunctionAppDomain/api/health"
}

# ============================================================================
# STEP 5: BUILD AGENT PACKAGE
# ============================================================================

Write-Information ""
Write-Information "  [5/5] Building agent package (.zip)..."

$packageResult = & (Join-Path $ScriptDir "Build-AgentPackage.ps1") -FromEntraOutput

# ============================================================================
# DONE
# ============================================================================

$packagePath = $packageResult.PackagePath

Write-Information ""
Write-Information "  ╔══════════════════════════════════════════════════════════╗"
Write-Information "  ║              DEPLOYMENT COMPLETE                         ║"
Write-Information "  ╠══════════════════════════════════════════════════════════╣"
Write-Information "  ║                                                          ║"
Write-Information "  ║  Function App:  https://$FunctionAppDomain"
Write-Information "  ║  Client ID:     $ClientId"
Write-Information "  ║  Tenant ID:     $TenantId"
Write-Information "  ║  Package:       $packagePath"
Write-Information "  ║                                                          ║"
Write-Information "  ║  ONE REMAINING STEP:                                    ║"
Write-Information "  ║                                                          ║"
Write-Information "  ║  Upload the .zip to Teams Admin Center:                 ║"
Write-Information "  ║  https://admin.teams.microsoft.com                      ║"
Write-Information "  ║  → Teams apps → Manage apps → Upload new app            ║"
Write-Information "  ║  → Select: $packagePath"
Write-Information "  ║  → Assign to users/groups as needed                     ║"
Write-Information "  ║                                                          ║"
Write-Information "  ╚══════════════════════════════════════════════════════════╝"
Write-Information ""

# Save full deployment record
$fullResult = @{
    AppName       = $AppName
    FunctionUrl   = "https://$FunctionAppDomain"
    ClientId      = $ClientId
    TenantId      = $TenantId
    TeamsAppId    = $packageResult.TeamsAppId
    PackagePath   = $packagePath
    Region        = $Region
    DeployedAt    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
}
$fullResult | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $ScriptDir "deployment-record.json")

return $fullResult
