<#
.SYNOPSIS
    Removes test data created by Setup-TestEnvironment.ps1.

.DESCRIPTION
    Removes document libraries, lists, and SharePoint groups created for testing.
    Does NOT remove external users from the site (those should be managed manually).

.NOTES
    Run Connect-TestSite.ps1 first to establish a connection.

.EXAMPLE
    .\Connect-TestSite.ps1
    .\Remove-TestEnvironment.ps1
#>

param(
    [switch]$Force
)

Write-Host "`n========================================" -ForegroundColor Red
Write-Host "  External Sharing Risk Scanner - Cleanup" -ForegroundColor Red
Write-Host "========================================`n" -ForegroundColor Red

# Verify connection
try {
    $web = Get-PnPWeb
    Write-Host "Connected to: $($web.Title) ($($web.Url))" -ForegroundColor Green
}
catch {
    Write-Host "ERROR: Not connected. Run Connect-TestSite.ps1 first." -ForegroundColor Red
    exit 1
}

if (-not $Force) {
    $confirm = Read-Host "This will DELETE test libraries, lists, and groups. Continue? (y/N)"
    if ($confirm -ne 'y' -and $confirm -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        exit 0
    }
}

# Remove libraries
$libraries = @("Contracts", "Project Files", "Shared Reports", "Internal Only")
Write-Host "`n--- Removing Document Libraries ---" -ForegroundColor Yellow
foreach ($lib in $libraries) {
    try {
        $existing = Get-PnPList -Identity $lib -ErrorAction SilentlyContinue
        if ($existing) {
            Remove-PnPList -Identity $lib -Force -ErrorAction Stop
            Write-Host "  [REMOVED] $lib" -ForegroundColor Green
        }
        else {
            Write-Host "  [NOT FOUND] $lib" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  [ERROR] Could not remove $lib : $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Remove custom list
Write-Host "`n--- Removing Lists ---" -ForegroundColor Yellow
try {
    $existing = Get-PnPList -Identity "Risk Tracking" -ErrorAction SilentlyContinue
    if ($existing) {
        Remove-PnPList -Identity "Risk Tracking" -Force -ErrorAction Stop
        Write-Host "  [REMOVED] Risk Tracking" -ForegroundColor Green
    }
    else {
        Write-Host "  [NOT FOUND] Risk Tracking" -ForegroundColor Gray
    }
}
catch {
    Write-Host "  [ERROR] Could not remove Risk Tracking: $($_.Exception.Message)" -ForegroundColor Red
}

# Remove groups
$groups = @("External Partners", "Project Reviewers", "Empty Test Group")
Write-Host "`n--- Removing SharePoint Groups ---" -ForegroundColor Yellow
foreach ($grp in $groups) {
    try {
        $existing = Get-PnPGroup -Identity $grp -ErrorAction SilentlyContinue
        if ($existing) {
            Remove-PnPGroup -Identity $grp -Force -ErrorAction Stop
            Write-Host "  [REMOVED] $grp" -ForegroundColor Green
        }
        else {
            Write-Host "  [NOT FOUND] $grp" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "  [ERROR] Could not remove $grp : $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Cleanup Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nNote: External users were NOT removed. Manage them manually in site settings." -ForegroundColor Yellow
