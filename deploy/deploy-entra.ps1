<#
.SYNOPSIS
    Configures Entra ID app registration for the Refresh Agent.
    Run AFTER the ARM/Bicep template has created the Azure Function App.

.DESCRIPTION
    Creates the Entra ID app registration with:
      - Files.ReadWrite and User.Read delegated permissions
      - Application ID URI with access_as_user scope
      - All 7 authorized client applications (Teams, M365, Outlook)
      - Client secret (24-month expiry)

    Then updates the Azure Function App settings with the Entra values.

.PARAMETER AppName
    The Function App name (must match what was used in the ARM deployment)

.PARAMETER ResourceGroupName
    Resource group name (default: rg-{AppName})

.EXAMPLE
    .\deploy-entra.ps1 -AppName "refresh-contoso"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AppName,

    [Parameter(Mandatory = $false)]
    [string]$ResourceGroupName = "rg-$AppName"
)

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"

$FunctionAppDomain = "$AppName.azurewebsites.net"
$EntraAppName = "Refresh Agent - $AppName"

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
# STEP 1: Connect to Microsoft Graph
# ============================================================================

Write-Information ""
Write-Information "  [1/5] Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes "Application.ReadWrite.All" -NoWelcome
$context = Get-MgContext
$TenantId = $context.TenantId
Write-Information "    ✓ Connected to tenant: $TenantId"

# ============================================================================
# STEP 2: Create App Registration
# ============================================================================

Write-Information "  [2/5] Creating Entra ID app registration..."

$existingApp = Get-MgApplication -Filter "displayName eq '$EntraAppName'" -ErrorAction SilentlyContinue
if ($existingApp) {
    Write-Information "    ⚠ App '$EntraAppName' already exists. Using existing."
    $app = $existingApp
}
else {
    $app = New-MgApplication -DisplayName $EntraAppName -SignInAudience "AzureADMyOrg"
    Write-Information "    ✓ App registered: $($app.AppId)"
}

$ClientId = $app.AppId
$ObjectId = $app.Id

# Ensure service principal exists
$sp = Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -ErrorAction SilentlyContinue
if (-not $sp) {
    $sp = New-MgServicePrincipal -AppId $ClientId
}

# ============================================================================
# STEP 3: Configure permissions, scope, and authorized clients
# ============================================================================

Write-Information "  [3/5] Configuring API permissions and OAuth scope..."

$graphResourceId = "00000003-0000-0000-c000-000000000000"
$filesReadWrite = "5c28f0bf-8a70-41f1-8ee2-e11b8e0ee7c7"
$userRead = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"

$scopeId = [guid]::NewGuid().ToString()
$appIdUri = "api://$FunctionAppDomain/$ClientId"

$updateParams = @{
    IdentifierUris = @($appIdUri)
    RequiredResourceAccess = @(
        @{
            ResourceAppId  = $graphResourceId
            ResourceAccess = @(
                @{ Id = $filesReadWrite; Type = "Scope" },
                @{ Id = $userRead; Type = "Scope" }
            )
        }
    )
    Api = @{
        Oauth2PermissionScopes = @(
            @{
                AdminConsentDescription = "Allows M365 Copilot to call the Refresh API on behalf of the signed-in user"
                AdminConsentDisplayName = "Access Refresh API as user"
                Id                      = $scopeId
                IsEnabled               = $true
                Type                    = "User"
                UserConsentDescription  = "Allow Refresh Agent to access your OneDrive files"
                UserConsentDisplayName  = "Access Refresh API"
                Value                   = "access_as_user"
            }
        )
        PreAuthorizedApplications = @(
            foreach ($clientApp in $AuthorizedClientApps) {
                @{
                    AppId                  = $clientApp.Id
                    DelegatedPermissionIds = @($scopeId)
                }
            }
        )
    }
}

Update-MgApplication -ApplicationId $ObjectId @updateParams
Write-Information "    ✓ API permissions configured (Files.ReadWrite, User.Read)"
Write-Information "    ✓ Application ID URI: $appIdUri"
Write-Information "    ✓ access_as_user scope created"
Write-Information "    ✓ 7 authorized client applications added"

# Grant admin consent
Write-Information "    Granting admin consent..."
az ad app permission admin-consent --id $ClientId 2>&1 | Out-Null
Write-Information "    ✓ Admin consent granted"

# ============================================================================
# STEP 4: Create Client Secret
# ============================================================================

Write-Information "  [4/5] Creating client secret..."

$secret = Add-MgApplicationPassword -ApplicationId $ObjectId -PasswordCredential @{
    DisplayName = "Refresh Agent Auto-Deploy"
    EndDateTime = (Get-Date).AddMonths(24)
}
$ClientSecret = $secret.SecretText
Write-Information "    ✓ Client secret created (expires: $($secret.EndDateTime))"

# ============================================================================
# STEP 5: Update Function App Settings
# ============================================================================

Write-Information "  [5/5] Updating Function App with Entra ID values..."

az functionapp config appsettings set `
    --name $AppName `
    --resource-group $ResourceGroupName `
    --settings `
        "ENTRA_TENANT_ID=$TenantId" `
        "ENTRA_CLIENT_ID=$ClientId" `
        "ENTRA_CLIENT_SECRET=$ClientSecret" `
    --output none 2>&1

Write-Information "    ✓ Function App settings updated"

# ============================================================================
# SUMMARY
# ============================================================================

Write-Information ""
Write-Information "  ┌──────────────────────────────────────────────────────┐"
Write-Information "  │  ENTRA ID CONFIGURATION COMPLETE                     │"
Write-Information "  ├──────────────────────────────────────────────────────┤"
Write-Information "  │  Client ID:     $ClientId"
Write-Information "  │  Tenant ID:     $TenantId"
Write-Information "  │  App ID URI:    $appIdUri"
Write-Information "  │  Function App:  $AppName"
Write-Information "  │                                                      │"
Write-Information "  │  REMAINING STEPS:                                    │"
Write-Information "  │  1. Deploy Python code to Azure Function App         │"
Write-Information "  │  2. Upload agent .zip in Teams Admin Center          │"
Write-Information "  └──────────────────────────────────────────────────────┘"
Write-Information ""

# Export values
$result = @{
    TenantId     = $TenantId
    ClientId     = $ClientId
    AppIdUri     = $appIdUri
    FunctionApp  = $AppName
    Domain       = $FunctionAppDomain
}

$result | ConvertTo-Json | Set-Content -Path (Join-Path $PSScriptRoot "entra-output.json")
Write-Information "  Values saved to: deploy/entra-output.json"

return $result
