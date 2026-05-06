#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Builds a baseline set of Conditional Access policies for Microsoft 365.

.DESCRIPTION
    Creates the following in your tenant:
      - Disables Security Defaults
      - Security Group  : "CA-Exclusions"      (empty, break-glass)
      - Named Location  : "Allowed-Country-USA" (USA only)
      - 5 Conditional Access Policies (P1) + 2 Risky User Policies (P2 only)

    The script automatically detects your Entra ID license:
      - P1 : Creates 5 baseline policies, skips Risky Users policies
      - P2 : Creates all 7 policies including Risky Users (High Block + Med/Low MFA)

    IMPORTANT: Policies are created as Disabled. After reviewing Sign-in Logs
    for 7-14 days, manually set each policy to Report-Only, then On.

.NOTES
    Required Graph Permissions (Delegated):
      - Policy.Read.All
      - Policy.ReadWrite.ConditionalAccess
      - Policy.ReadWrite.AuthenticationMethod
      - Group.ReadWrite.All
      - Directory.Read.All
      - Application.Read.All

    Required Module:
      Install-Module Microsoft.Graph -Scope CurrentUser

    Author  : Christopher Johnston (christopher.johnston@it1.com)
    Version : 2.3  — Added P1/P2 license detection, Risky Users policies
                     created automatically when P2 is detected
#>

[CmdletBinding()]
param()

#region ── CONNECT ──────────────────────────────────────────────────────────────

Write-Host "`n📡  Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes `
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess",
    "Policy.ReadWrite.AuthenticationMethod",
    "Group.ReadWrite.All",
    "Directory.Read.All",
    "Application.Read.All" `
    -NoWelcome

Write-Host "✅  Connected.`n" -ForegroundColor Green

#endregion

#region ── HELPERS ──────────────────────────────────────────────────────────────

function Write-Step { param([string]$M); Write-Host "  ➤  $M" -ForegroundColor Yellow }
function Write-Done { param([string]$M); Write-Host "  ✅  $M`n" -ForegroundColor Green }
function Write-Warn { param([string]$M); Write-Host "  ⚠️   $M`n" -ForegroundColor DarkYellow }
function Write-Fail { param([string]$M); Write-Host "  ❌  $M`n" -ForegroundColor Red }

function New-CAPolicy {
    param([string]$DisplayName, [string]$JsonBody)
    try {
        $all      = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction Stop
        $existing = $all.value | Where-Object { $_.displayName -eq $DisplayName }
        if ($existing) {
            Write-Warn "Policy '$DisplayName' already exists — skipping."
            $script:policyResults.Add([PSCustomObject]@{ Name = $DisplayName; Status = "Skipped (exists)" })
            return
        }
        Write-Step "Creating: '$DisplayName'..."
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" `
            -Body $JsonBody -ContentType "application/json" -ErrorAction Stop | Out-Null
        Write-Done "'$DisplayName' created."
        $script:policyResults.Add([PSCustomObject]@{ Name = $DisplayName; Status = "Created ✅" })
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Fail "Failed '$DisplayName': $_"
        $script:policyResults.Add([PSCustomObject]@{ Name = $DisplayName; Status = "FAILED ❌" })
    }
}

#endregion

#region ── METADATA ─────────────────────────────────────────────────────────────

$script:policyCreator    = "Christopher Johnston"
$script:policyCreatorUPN = "christopher.johnston@it1.com"
$script:policyCreatedOn  = (Get-Date -Format "yyyy-MM-dd")
$script:policyResults    = [System.Collections.Generic.List[PSCustomObject]]::new()

#endregion

#region ── STEP 0 › SECURITY DEFAULTS ───────────────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " STEP 0 › Checking Security Defaults"               -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

try {
    Write-Step "Checking current Security Defaults state..."
    $secDefaults = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" -ErrorAction Stop
    if ($secDefaults.isEnabled -eq $true) {
        Write-Host "  ⚠️   Security Defaults is ENABLED — must be disabled first.`n" -ForegroundColor DarkYellow
        $confirm = Read-Host "  Disable Security Defaults now? (yes/no)"
        if ($confirm -eq "yes") {
            Write-Step "Disabling Security Defaults..."
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" `
                -Body '{"isEnabled":false}' -ContentType "application/json" -ErrorAction Stop | Out-Null
            Write-Done "Security Defaults DISABLED."
            Write-Host "  ⚠️   Tenant has NO enforced MFA until CA policies are turned On.`n" -ForegroundColor DarkYellow
        } else {
            Write-Fail "Aborted. Disable Security Defaults manually then re-run."
            Disconnect-MgGraph | Out-Null; exit 1
        }
    } else {
        Write-Done "Security Defaults already DISABLED. Proceeding."
    }
} catch {
    Write-Fail "Could not check Security Defaults: $_"
    Disconnect-MgGraph | Out-Null; exit 1
}

#endregion

#region ── STEP 0.5 › LICENSE CHECK ─────────────────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " STEP 0.5 › Checking Entra ID License (P1 vs P2)"  -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

try {
    Write-Step "Checking subscribed SKUs..."
    $skus = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -ErrorAction Stop

    # SKU part numbers that indicate P2
    $p2SkuPartNumbers = @(
        "AAD_PREMIUM_P2",           # Entra ID P2 standalone
        "ENTERPRISEPREMIUM",        # M365 E5
        "SPE_E5",                   # Microsoft 365 E5
        "IDENTITY_THREAT_PROTECTION", # Microsoft Entra ID P2 (new name)
        "AAD_PREMIUM_P2_FACULTY",
        "AAD_PREMIUM_P2_STUDENT"
    )

    $script:hasP2 = $false
    foreach ($sku in $skus.value) {
        if ($sku.capabilityStatus -eq "Enabled" -and $p2SkuPartNumbers -contains $sku.skuPartNumber) {
            $script:hasP2 = $true
            Write-Host "  ✅  Entra ID P2 detected: $($sku.skuPartNumber)`n" -ForegroundColor Green
            break
        }
    }

    if (-not $script:hasP2) {
        Write-Host "  ℹ️   No Entra ID P2 license detected." -ForegroundColor Cyan
        Write-Host "       Risky Users policies will be skipped (require P2).`n" -ForegroundColor Cyan
    }
} catch {
    Write-Warn "Could not check license. Defaulting to P1 — Risky Users policies will be skipped."
    $script:hasP2 = $false
}

#endregion

#region ── STEP 1 › CA-EXCLUSIONS GROUP ─────────────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " STEP 1 › Creating CA-Exclusions Security Group" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

try {
    Write-Step "Checking if group 'CA-Exclusions' already exists..."
    $groupResult = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq 'CA-Exclusions'&`$top=1" -ErrorAction Stop
    if ($groupResult.value.Count -gt 0) {
        $script:exclusionGroupId = $groupResult.value[0].id
        Write-Warn "Group already exists. Using ID: $($script:exclusionGroupId)"
    } else {
        Write-Step "Creating group 'CA-Exclusions'..."
        $newGroup = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/groups" `
            -Body '{"displayName":"CA-Exclusions","description":"Break-glass exclusion group for Conditional Access policies. Add accounts here only when required.","mailEnabled":false,"mailNickname":"CA-Exclusions","securityEnabled":true}' `
            -ContentType "application/json" -ErrorAction Stop
        $script:exclusionGroupId = $newGroup.id
        Write-Done "Group created. ID: $($script:exclusionGroupId)"
    }
} catch {
    Write-Fail "Failed to create/retrieve group: $_"
    Disconnect-MgGraph | Out-Null; exit 1
}

#endregion

#region ── STEP 2 › NAMED LOCATION ──────────────────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " STEP 2 › Creating Named Location (USA Only)"   -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

try {
    Write-Step "Checking if 'Allowed-Country-USA' already exists..."
    $locationResult = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations" -ErrorAction Stop
    $existingLocation = $locationResult.value | Where-Object { $_.displayName -eq "Allowed-Country-USA" }
    if ($existingLocation) {
        $script:namedLocationId = $existingLocation.id
        Write-Warn "Named location already exists. Using ID: $($script:namedLocationId)"
    } else {
        Write-Step "Creating named location 'Allowed-Country-USA'..."
        $newLocation = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations" `
            -Body '{"@odata.type":"#microsoft.graph.countryNamedLocation","displayName":"Allowed-Country-USA","countriesAndRegions":["US"],"includeUnknownCountriesAndRegions":false}' `
            -ContentType "application/json" -ErrorAction Stop
        $script:namedLocationId = $newLocation.id
        Write-Done "Named location created. ID: $($script:namedLocationId)"
        Write-Step "Waiting 15 seconds for named location to replicate..."
        Start-Sleep -Seconds 15
        Write-Host "  ✅  Replication wait complete.`n" -ForegroundColor Green
    }
} catch {
    Write-Fail "Failed to create/retrieve named location: $_"
    Disconnect-MgGraph | Out-Null; exit 1
}

#endregion

#region ── STEP 3 › CONDITIONAL ACCESS POLICIES ─────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " STEP 3 › Creating Conditional Access Policies"               -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

Write-Host "  Using Group ID    : $($script:exclusionGroupId)" -ForegroundColor DarkGray
Write-Host "  Using Location ID : $($script:namedLocationId)" -ForegroundColor DarkGray
Write-Host "  License Tier      : $(if ($script:hasP2) { 'P2 — All 7 policies will be created' } else { 'P1 — Risky Users policies will be skipped' })`n" -ForegroundColor DarkGray

if ([string]::IsNullOrEmpty($script:exclusionGroupId)) {
    Write-Fail "exclusionGroupId is empty — cannot create policies. Exiting."
    Disconnect-MgGraph | Out-Null; exit 1
}
if ([string]::IsNullOrEmpty($script:namedLocationId)) {
    Write-Fail "namedLocationId is empty — cannot create policies. Exiting."
    Disconnect-MgGraph | Out-Null; exit 1
}

# ── POLICY 1 › Block Legacy Authentication ───────────────────────────────────
New-CAPolicy -DisplayName "BASELINE - Blocking Legacy Authentication" -JsonBody (
    '{"displayName":"BASELINE - Blocking Legacy Authentication",' +
    '"state":"disabled",' +
    '"conditions":{' +
        '"users":{"includeUsers":["All"],"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
        '"applications":{"includeApplications":["All"]},' +
        '"clientAppTypes":["exchangeActiveSync","other"]},' +
    '"grantControls":{"operator":"OR","builtInControls":["block"]}}'
)

# ── POLICY 2 › GEO Block — Only USA Allowed ──────────────────────────────────
New-CAPolicy -DisplayName "BASELINE - GEO Block (Only USA Allowed)" -JsonBody (
    '{"displayName":"BASELINE - GEO Block (Only USA Allowed)",' +
    '"state":"disabled",' +
    '"conditions":{' +
        '"users":{"includeUsers":["All"],"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
        '"applications":{"includeApplications":["All"]},' +
        '"locations":{"includeLocations":["All"],"excludeLocations":["' + $script:namedLocationId + '"]}},' +
    '"grantControls":{"operator":"OR","builtInControls":["block"]}}'
)

# ── POLICY 3 › Require MFA for All Users ─────────────────────────────────────
New-CAPolicy -DisplayName "BASELINE - MFA All Users" -JsonBody (
    '{"displayName":"BASELINE - MFA All Users",' +
    '"state":"disabled",' +
    '"conditions":{' +
        '"users":{"includeUsers":["All"],"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
        '"applications":{"includeApplications":["All"]},' +
        '"clientAppTypes":["all"]},' +
    '"grantControls":{"operator":"OR","builtInControls":["mfa"]}}'
)

# ── POLICY 4 › Only Allow Windows, Mac, iOS, Android ─────────────────────────
New-CAPolicy -DisplayName "BASELINE - Only Allow Windows, Mac, iOS, Android" -JsonBody (
    '{"displayName":"BASELINE - Only Allow Windows, Mac, iOS, Android",' +
    '"state":"disabled",' +
    '"conditions":{' +
        '"users":{"includeUsers":["All"],"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
        '"applications":{"includeApplications":["All"]},' +
        '"platforms":{"includePlatforms":["all"],"excludePlatforms":["windows","macOS","iOS","android"]}},' +
    '"grantControls":{"operator":"OR","builtInControls":["block"]}}'
)

# ── POLICY 5 › Phishing-Resistant MFA for Admin Portals ──────────────────────
New-CAPolicy -DisplayName "BASELINE - Phishing-Resistant MFA for Admin Portals" -JsonBody (
    '{"displayName":"BASELINE - Phishing-Resistant MFA for Admin Portals",' +
    '"state":"disabled",' +
    '"conditions":{' +
        '"users":{"includeRoles":[' +
            '"62e90394-69f5-4237-9190-012177145e10",' +
            '"194ae4cb-b126-40b2-bd5b-6091b380977d",' +
            '"f28a1f50-f6e7-4571-818b-6a12f2af6b6c",' +
            '"29232cdf-9323-42fd-ade2-1d097af3e4de",' +
            '"b1be1c3e-b65d-4f19-8427-f6fa0d97feb9",' +
            '"9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3",' +
            '"158c047a-c907-4556-b7ef-446551a6b5f7",' +
            '"7be44c8a-adaf-4e2a-84d6-ab2649e08a13",' +
            '"e8611ab8-c189-46e8-94e1-60213ab1f814",' +
            '"966707d0-3269-4727-9be2-8c3a10f19b9d",' +
            '"f70938a0-fc10-4177-9e90-2178f8765737",' +
            '"69091246-20e8-4a56-aa4d-066075b2a7a8",' +
            '"4d6ac14f-3453-41d0-bef9-a3e0c569773a",' +
            '"fe930be7-5e62-47db-91af-98c3a49a38b1"],' +
        '"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
        '"applications":{"includeApplications":["MicrosoftAdminPortals"]},' +
        '"clientAppTypes":["all"]},' +
    '"grantControls":{"operator":"OR","authenticationStrength":{"id":"00000000-0000-0000-0000-000000000004"}}}'
)

# ── POLICY 6 › Risky Users — High Risk (Block) — P2 ONLY ─────────────────────
if ($script:hasP2) {
    New-CAPolicy -DisplayName "BASELINE - Risky Users Risk - High (Block)" -JsonBody (
        '{"displayName":"BASELINE - Risky Users Risk - High (Block)",' +
        '"state":"disabled",' +
        '"conditions":{' +
            '"users":{"includeUsers":["All"],"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
            '"applications":{"includeApplications":["All"]},' +
            '"userRiskLevels":["high"]},' +
        '"grantControls":{"operator":"OR","builtInControls":["block"]}}'
    )
} else {
    Write-Host "  ⏭️   Skipping 'BASELINE - Risky Users Risk - High (Block)' — requires P2.`n" -ForegroundColor DarkGray
    $script:policyResults.Add([PSCustomObject]@{ Name = "BASELINE - Risky Users Risk - High (Block)"; Status = "Skipped (P1 license)" })
}

# ── POLICY 7 › Risky Users — Med/Low Risk (MFA + Password Change) — P2 ONLY ──
if ($script:hasP2) {
    New-CAPolicy -DisplayName "BASELINE - Risky Users Risk Med-Low (Enforce MFA)" -JsonBody (
        '{"displayName":"BASELINE - Risky Users Risk Med-Low (Enforce MFA)",' +
        '"state":"disabled",' +
        '"conditions":{' +
            '"users":{"includeUsers":["All"],"excludeUsers":[],"excludeGroups":["' + $script:exclusionGroupId + '"]},' +
            '"applications":{"includeApplications":["All"]},' +
            '"userRiskLevels":["medium","low"]},' +
        '"grantControls":{"operator":"AND","builtInControls":["mfa","passwordChange"]}}'
    )
} else {
    Write-Host "  ⏭️   Skipping 'BASELINE - Risky Users Risk Med-Low (Enforce MFA)' — requires P2.`n" -ForegroundColor DarkGray
    $script:policyResults.Add([PSCustomObject]@{ Name = "BASELINE - Risky Users Risk Med-Low (Enforce MFA)"; Status = "Skipped (P1 license)" })
}

#endregion

#region ── SUMMARY ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " ✅  BASELINE DEPLOYMENT COMPLETE"                             -ForegroundColor Green
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " Deployed by  : $($script:policyCreator) ($($script:policyCreatorUPN))" -ForegroundColor Gray
Write-Host " Deployed on  : $($script:policyCreatedOn)"                    -ForegroundColor Gray
Write-Host " License tier : $(if ($script:hasP2) { 'Entra ID P2 — all 7 policies deployed' } else { 'Entra ID P1 — 5 policies deployed, 2 skipped' })" -ForegroundColor Gray
Write-Host ""
Write-Host " Resources:"                                                   -ForegroundColor White
Write-Host "   • Security Defaults  : Disabled"                            -ForegroundColor Gray
Write-Host "   • Security Group     : CA-Exclusions      ($($script:exclusionGroupId))" -ForegroundColor Gray
Write-Host "   • Named Location     : Allowed-Country-USA ($($script:namedLocationId))" -ForegroundColor Gray
Write-Host ""
Write-Host " Policy Results:"                                              -ForegroundColor White
foreach ($result in $script:policyResults) {
    Write-Host ("   {0,-58} {1}" -f $result.Name, $result.Status)         -ForegroundColor Gray
}
Write-Host ""
Write-Host " ⚠️   IMPORTANT: All policies created as DISABLED."            -ForegroundColor DarkYellow
Write-Host "       Enable in this order when ready:"                       -ForegroundColor DarkYellow
Write-Host "   1) Blocking Legacy Authentication"                          -ForegroundColor DarkYellow
Write-Host "   2) MFA All Users"                                           -ForegroundColor DarkYellow
Write-Host "   3) Only Allow Windows, Mac, iOS, Android"                   -ForegroundColor DarkYellow
Write-Host "   4) Phishing-Resistant MFA for Admin Portals"                -ForegroundColor DarkYellow
if ($script:hasP2) {
Write-Host "   5) Risky Users - High (Block)"                              -ForegroundColor DarkYellow
Write-Host "   6) Risky Users Med-Low (Enforce MFA)"                       -ForegroundColor DarkYellow
Write-Host "   7) GEO Block (Only USA Allowed)  ← LAST"                   -ForegroundColor DarkYellow
} else {
Write-Host "   5) GEO Block (Only USA Allowed)  ← LAST"                   -ForegroundColor DarkYellow
Write-Host ""
Write-Host "   ℹ️   Upgrade to Entra ID P2 to deploy Risky Users policies." -ForegroundColor Cyan
}
Write-Host ""
Write-Host " Portal: https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

Disconnect-MgGraph | Out-Null

#endregion
