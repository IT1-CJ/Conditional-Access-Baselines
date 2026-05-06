#Requires -Modules Microsoft.Graph.Authentication

<#
.SYNOPSIS
    Removes all Conditional Access policies and named locations from the tenant.

.DESCRIPTION
    This script connects to Microsoft Graph and permanently deletes:
      - All Conditional Access policies
      - All Named Locations

    Use this script to clean up a tenant before re-running the baseline
    deployment script, or to reset CA policies entirely.

    WARNING: This is a destructive operation. All CA policies and named
    locations will be permanently deleted. There is no undo.

.NOTES
    Required Graph Permissions (Delegated):
      - Policy.ReadWrite.ConditionalAccess
      - Directory.Read.All

    Required Module:
      Install-Module Microsoft.Graph -Scope CurrentUser

    Author  : Christopher Johnston (christopher.johnston@it1.com)
    Version : 1.0
#>

[CmdletBinding()]
param(
    [switch]$Force  # Skip confirmation prompt
)

#region ── CONNECT ──────────────────────────────────────────────────────────────

Write-Host "`n📡  Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes `
    "Policy.ReadWrite.ConditionalAccess",
    "Directory.Read.All" `
    -NoWelcome

Write-Host "✅  Connected.`n" -ForegroundColor Green

#endregion

#region ── INVENTORY ────────────────────────────────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " Scanning tenant for CA policies and named locations" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

try {
    $policies  = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction Stop
    $locations = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations" -ErrorAction Stop
} catch {
    Write-Host "  ❌  Failed to retrieve policies/locations: $_" -ForegroundColor Red
    Disconnect-MgGraph | Out-Null; exit 1
}

$policyCount   = $policies.value.Count
$locationCount = $locations.value.Count

if ($policyCount -eq 0 -and $locationCount -eq 0) {
    Write-Host "  ✅  No CA policies or named locations found. Nothing to delete.`n" -ForegroundColor Green
    Disconnect-MgGraph | Out-Null; exit 0
}

Write-Host "  Found $policyCount CA polic$(if ($policyCount -eq 1) {'y'} else {'ies'}):" -ForegroundColor White
foreach ($policy in $policies.value) {
    Write-Host "    • $($policy.displayName)  [$($policy.state)]" -ForegroundColor Gray
}

Write-Host ""
Write-Host "  Found $locationCount named location$(if ($locationCount -eq 1) {''} else {'s'}):" -ForegroundColor White
foreach ($location in $locations.value) {
    Write-Host "    • $($location.displayName)" -ForegroundColor Gray
}

#endregion

#region ── CONFIRM ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " ⚠️   WARNING: This will permanently delete everything" -ForegroundColor Red
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

if (-not $Force) {
    $confirm = Read-Host "  Type 'DELETE' to confirm permanent deletion of all CA policies and named locations"
    if ($confirm -ne "DELETE") {
        Write-Host "`n  Aborted. Nothing was deleted.`n" -ForegroundColor Cyan
        Disconnect-MgGraph | Out-Null; exit 0
    }
}

#endregion

#region ── DELETE POLICIES ──────────────────────────────────────────────────────

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " Deleting CA Policies"                               -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

$deletedPolicies  = 0
$failedPolicies   = 0

foreach ($policy in $policies.value) {
    try {
        Write-Host "  ➤  Deleting policy: $($policy.displayName)..." -ForegroundColor Yellow
        Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($policy.id)" `
            -ErrorAction Stop | Out-Null
        Write-Host "  ✅  Deleted.`n" -ForegroundColor Green
        $deletedPolicies++
    } catch {
        Write-Host "  ❌  Failed to delete '$($policy.displayName)': $_`n" -ForegroundColor Red
        $failedPolicies++
    }
}

#endregion

#region ── DELETE NAMED LOCATIONS ───────────────────────────────────────────────

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " Deleting Named Locations"                           -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

# Wait briefly to allow policy deletions to replicate before attempting location deletes
if ($locationCount -gt 0 -and $deletedPolicies -gt 0) {
    Write-Host "  ➤  Waiting 10 seconds for policy deletions to replicate..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10
    Write-Host "  ✅  Replication wait complete.`n" -ForegroundColor Green
}

$deletedLocations = 0
$failedLocations  = 0

foreach ($location in $locations.value) {
    try {
        Write-Host "  ➤  Deleting location: $($location.displayName)..." -ForegroundColor Yellow
        Invoke-MgGraphRequest -Method DELETE `
            -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/$($location.id)" `
            -ErrorAction Stop | Out-Null
        Write-Host "  ✅  Deleted.`n" -ForegroundColor Green
        $deletedLocations++
    } catch {
        Write-Host "  ❌  Failed to delete '$($location.displayName)': $_`n" -ForegroundColor Red
        Write-Host "      If the error is 'referenced by policies', wait a moment and re-run the script." -ForegroundColor Gray
        $failedLocations++
    }
}

#endregion

#region ── SUMMARY ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host " ✅  CLEANUP COMPLETE"                                         -ForegroundColor Green
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host "   Policies deleted       : $deletedPolicies"                  -ForegroundColor Gray
Write-Host "   Policies failed        : $failedPolicies"                   -ForegroundColor Gray
Write-Host "   Named locations deleted: $deletedLocations"                 -ForegroundColor Gray
Write-Host "   Named locations failed : $failedLocations"                  -ForegroundColor Gray

if ($failedPolicies -gt 0 -or $failedLocations -gt 0) {
    Write-Host ""
    Write-Host "  ⚠️   Some items failed to delete. Re-run the script to retry." -ForegroundColor DarkYellow
}

Write-Host ""
Write-Host " Portal: https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray

Disconnect-MgGraph | Out-Null

#endregion
