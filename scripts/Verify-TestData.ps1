<#
.SYNOPSIS
    Verifies that test data is in place for the External Sharing Risk Scanner.

.DESCRIPTION
    Checks for the expected libraries, lists, groups, and permission configurations.
    Use this before each test run to confirm the environment is ready.

.NOTES
    Run Connect-TestSite.ps1 first to establish a connection.

.EXAMPLE
    .\Connect-TestSite.ps1
    .\Verify-TestData.ps1
#>

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  External Sharing Risk Scanner - Verify" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Verify connection
try {
    $web = Get-PnPWeb
    Write-Host "Connected to: $($web.Title) ($($web.Url))" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Not connected. Run Connect-TestSite.ps1 first." -ForegroundColor Red
    exit 1
}

$passed = 0
$failed = 0
$warnings = 0

function Test-Check {
    param([string]$Name, [bool]$Result)
    if ($Result) {
        Write-Host "  ✓ $Name" -ForegroundColor Green
        $script:passed++
    }
    else {
        Write-Host "  ✗ $Name" -ForegroundColor Red
        $script:failed++
    }
}

function Test-Warning {
    param([string]$Message)
    Write-Host "  ⚠ $Message" -ForegroundColor Yellow
    $script:warnings++
}

# ============================================================
# CHECK LIBRARIES
# ============================================================
Write-Host "`n--- Document Libraries ---" -ForegroundColor Yellow

$expectedLibraries = @("Contracts", "Project Files", "Shared Reports", "Internal Only")
foreach ($lib in $expectedLibraries) {
    $exists = $null -ne (Get-PnPList -Identity $lib -ErrorAction SilentlyContinue)
    Test-Check "Library: $lib" $exists
}

# Check unique permissions
$contracts = Get-PnPList -Identity "Contracts" -ErrorAction SilentlyContinue
if ($contracts) {
    Test-Check "Contracts has unique permissions" $contracts.HasUniqueRoleAssignments
}

$sharedReports = Get-PnPList -Identity "Shared Reports" -ErrorAction SilentlyContinue
if ($sharedReports) {
    Test-Check "Shared Reports has unique permissions" $sharedReports.HasUniqueRoleAssignments
}

$projectFiles = Get-PnPList -Identity "Project Files" -ErrorAction SilentlyContinue
if ($projectFiles) {
    Test-Check "Project Files inherits permissions" (-not $projectFiles.HasUniqueRoleAssignments)
}

# ============================================================
# CHECK LISTS
# ============================================================
Write-Host "`n--- Custom Lists ---" -ForegroundColor Yellow
$riskTracking = Get-PnPList -Identity "Risk Tracking" -ErrorAction SilentlyContinue
Test-Check "List: Risk Tracking" ($null -ne $riskTracking)

# ============================================================
# CHECK GROUPS
# ============================================================
Write-Host "`n--- SharePoint Groups ---" -ForegroundColor Yellow

$expectedGroups = @("External Partners", "Project Reviewers", "Empty Test Group")
foreach ($grp in $expectedGroups) {
    $exists = $null -ne (Get-PnPGroup -Identity $grp -ErrorAction SilentlyContinue)
    Test-Check "Group: $grp" $exists
}

# Check empty group
$emptyGroup = Get-PnPGroup -Identity "Empty Test Group" -ErrorAction SilentlyContinue
if ($emptyGroup) {
    $members = Get-PnPGroupMember -Identity "Empty Test Group" -ErrorAction SilentlyContinue
    $isEmpty = ($null -eq $members) -or ($members.Count -eq 0)
    Test-Check "Empty Test Group is empty" $isEmpty
}

# ============================================================
# CHECK EXTERNAL USERS
# ============================================================
Write-Host "`n--- External Users ---" -ForegroundColor Yellow

$siteUsers = Get-PnPUser -ErrorAction SilentlyContinue
$externalUsers = $siteUsers | Where-Object { $_.LoginName -like "*#ext#*" }

if ($externalUsers -and $externalUsers.Count -gt 0) {
    Write-Host "  ✓ Found $($externalUsers.Count) external user(s):" -ForegroundColor Green
    foreach ($ext in $externalUsers) {
        Write-Host "    - $($ext.Title) ($($ext.Email))" -ForegroundColor Gray
    }
    $passed++
}
else {
    Test-Warning "No external users found. Invite guests to test external detection."
}

# ============================================================
# SUMMARY
# ============================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Verification Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Passed:   $passed" -ForegroundColor Green
Write-Host "  Failed:   $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Gray" })
Write-Host "  Warnings: $warnings" -ForegroundColor $(if ($warnings -gt 0) { "Yellow" } else { "Gray" })

if ($failed -gt 0) {
    Write-Host "`n  Run Setup-TestEnvironment.ps1 to create missing test data." -ForegroundColor Yellow
}
else {
    Write-Host "`n  Environment is ready for testing! 🎉" -ForegroundColor Green
}
