<#
.SYNOPSIS
    Builds the Refresh Agent app package (.zip) ready for sideloading into a customer tenant.

.DESCRIPTION
    Takes the appPackage/ template files, stamps in the customer's specific values
    (Function App URL, Entra client ID, tenant ID), and produces a .zip file that
    can be uploaded via Teams Admin Center.

    The .zip contains: manifest.json, declarativeAgent.json, ai-plugin.json,
    openapi.yaml, adaptive cards, and icons — everything needed for the agent.

.PARAMETER AppName
    The Function App name (used to derive the domain)

.PARAMETER EntraClientId
    The customer's Entra ID Application (client) ID

.PARAMETER EntraTenantId
    The customer's Entra ID tenant ID (GUID)

.PARAMETER TeamsAppId
    Teams App ID (GUID). If not provided, a new GUID is generated.

.PARAMETER OutputPath
    Path for the output .zip file (default: deploy/refresh-agent-package.zip)

.PARAMETER FromEntraOutput
    If specified, reads values from deploy/entra-output.json instead of parameters

.EXAMPLE
    # Using explicit parameters:
    .\Build-AgentPackage.ps1 `
        -AppName "refresh-contoso" `
        -EntraClientId "a1b2c3d4-..." `
        -EntraTenantId "e5f6g7h8-..."

    # Using output from deploy-entra.ps1:
    .\Build-AgentPackage.ps1 -FromEntraOutput
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppName,

    [Parameter(Mandatory = $false)]
    [string]$EntraClientId,

    [Parameter(Mandatory = $false)]
    [string]$EntraTenantId,

    [Parameter(Mandatory = $false)]
    [string]$TeamsAppId = [guid]::NewGuid().ToString(),

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",

    [switch]$FromEntraOutput
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$ScriptDir = $PSScriptRoot
$ProjectRoot = Split-Path $ScriptDir -Parent

# ============================================================================
# LOAD VALUES
# ============================================================================

if ($FromEntraOutput) {
    $entraOutputPath = Join-Path $ScriptDir "entra-output.json"
    if (-not (Test-Path $entraOutputPath)) {
        throw "entra-output.json not found. Run deploy-entra.ps1 first."
    }
    $entraValues = Get-Content $entraOutputPath | ConvertFrom-Json
    $AppName = $entraValues.FunctionApp
    $EntraClientId = $entraValues.ClientId
    $EntraTenantId = $entraValues.TenantId
    Write-Information "  Loaded values from entra-output.json"
}

if (-not $AppName -or -not $EntraClientId -or -not $EntraTenantId) {
    throw "AppName, EntraClientId, and EntraTenantId are required. Provide them as parameters or use -FromEntraOutput."
}

$FunctionAppDomain = "$AppName.azurewebsites.net"
$FunctionAppUrl = "https://$FunctionAppDomain"

if (-not $OutputPath) {
    $OutputPath = Join-Path $ScriptDir "refresh-agent-$AppName.zip"
}

# ============================================================================
# BUILD PACKAGE
# ============================================================================

Write-Information ""
Write-Information "  Building Refresh Agent package for: $AppName"
Write-Information "  ─────────────────────────────────────────────"

# Create temp directory
$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) "refresh-agent-build-$(Get-Random)"
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

$appPackageSrc = Join-Path $ProjectRoot "appPackage"

try {
    # --- Copy all files to temp directory ---
    Write-Information "  [1/4] Copying template files..."

    # Top-level files
    foreach ($file in @("declarativeAgent.json", "ai-plugin.json", "color.png", "outline.png")) {
        $srcPath = Join-Path $appPackageSrc $file
        if (Test-Path $srcPath) {
            Copy-Item $srcPath -Destination $tempDir
        }
        else {
            Write-Warning "    Missing: $file"
        }
    }

    # apiSpecificationFile
    $apiSpecDir = Join-Path $tempDir "apiSpecificationFile"
    New-Item -ItemType Directory -Path $apiSpecDir -Force | Out-Null
    Copy-Item (Join-Path $appPackageSrc "apiSpecificationFile" "openapi.yaml") -Destination $apiSpecDir

    # adaptiveCards
    $cardsDir = Join-Path $tempDir "adaptiveCards"
    New-Item -ItemType Directory -Path $cardsDir -Force | Out-Null
    Get-ChildItem (Join-Path $appPackageSrc "adaptiveCards" "*.json") | ForEach-Object {
        Copy-Item $_.FullName -Destination $cardsDir
    }

    Write-Information "    ✓ All template files copied"

    # --- Stamp manifest.json ---
    Write-Information "  [2/4] Stamping manifest.json..."

    $manifestSrc = Get-Content (Join-Path $appPackageSrc "manifest.json") -Raw
    $manifestSrc = $manifestSrc -replace '\$\{\{TEAMS_APP_ID\}\}', $TeamsAppId
    $manifestSrc = $manifestSrc -replace '\$\{\{AZURE_FUNCTION_DOMAIN\}\}', $FunctionAppDomain
    $manifestSrc = $manifestSrc -replace '\$\{\{ENTRA_CLIENT_ID\}\}', $EntraClientId
    $manifestSrc = $manifestSrc -replace '\$\{\{APP_NAME_SUFFIX\}\}', ''
    Set-Content -Path (Join-Path $tempDir "manifest.json") -Value $manifestSrc -NoNewline

    Write-Information "    ✓ Teams App ID:  $TeamsAppId"
    Write-Information "    ✓ Domain:        $FunctionAppDomain"
    Write-Information "    ✓ Client ID:     $EntraClientId"

    # --- Stamp openapi.yaml ---
    Write-Information "  [3/4] Stamping openapi.yaml..."

    $openapiPath = Join-Path $apiSpecDir "openapi.yaml"
    $openapiContent = Get-Content $openapiPath -Raw
    $openapiContent = $openapiContent -replace '\$\{\{AZURE_FUNCTION_URL\}\}', $FunctionAppUrl
    $openapiContent = $openapiContent -replace '\$\{\{ENTRA_TENANT_ID\}\}', $EntraTenantId
    Set-Content -Path $openapiPath -Value $openapiContent -NoNewline

    Write-Information "    ✓ Function URL:  $FunctionAppUrl"
    Write-Information "    ✓ Tenant ID:     $EntraTenantId"

    # Note: ai-plugin.json still has ${{OAUTH2_REGISTRATION_ID}} which gets
    # resolved at runtime by Teams/Copilot. For sideloaded packages, we
    # replace it with a placeholder reference that Teams handles.
    $pluginPath = Join-Path $tempDir "ai-plugin.json"
    $pluginContent = Get-Content $pluginPath -Raw
    # For sideloaded packages, use the Entra client ID as the OAuth reference
    $pluginContent = $pluginContent -replace '\$\{\{OAUTH2_REGISTRATION_ID\}\}', "$EntraClientId"
    Set-Content -Path $pluginPath -Value $pluginContent -NoNewline

    # --- Create .zip ---
    Write-Information "  [4/4] Creating .zip package..."

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    # Compress from temp directory
    Compress-Archive -Path "$tempDir\*" -DestinationPath $OutputPath -Force

    $zipSize = (Get-Item $OutputPath).Length / 1KB
    Write-Information "    ✓ Package created: $OutputPath ($([math]::Round($zipSize, 1)) KB)"
}
finally {
    # Cleanup temp directory
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
}

# ============================================================================
# SUMMARY
# ============================================================================

Write-Information ""
Write-Information "  ┌──────────────────────────────────────────────────────┐"
Write-Information "  │  AGENT PACKAGE READY                                 │"
Write-Information "  ├──────────────────────────────────────────────────────┤"
Write-Information "  │  Package:  $OutputPath"
Write-Information "  │  App ID:   $TeamsAppId"
Write-Information "  │  Target:   $FunctionAppUrl"
Write-Information "  │                                                      │"
Write-Information "  │  TO INSTALL:                                         │"
Write-Information "  │  1. Go to Teams Admin Center                        │"
Write-Information "  │     (https://admin.teams.microsoft.com)             │"
Write-Information "  │  2. Teams apps → Manage apps → Upload new app       │"
Write-Information "  │  3. Select this .zip file                           │"
Write-Information "  │  4. Assign the app to users/groups as needed        │"
Write-Information "  │                                                      │"
Write-Information "  │  OR for testing:                                     │"
Write-Information "  │  1. Go to Teams → Apps → Manage your apps           │"
Write-Information "  │  2. Upload a custom app → Select this .zip          │"
Write-Information "  └──────────────────────────────────────────────────────┘"
Write-Information ""

return @{
    PackagePath = $OutputPath
    TeamsAppId  = $TeamsAppId
    FunctionApp = $AppName
}
