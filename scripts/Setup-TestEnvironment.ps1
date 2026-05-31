<#
.SYNOPSIS
    Provisions test data for the External Sharing Risk Scanner on the lab site.

.DESCRIPTION
    Creates document libraries, lists, SharePoint groups, and simulates
    external user scenarios for testing the web part's detection capabilities.

.NOTES
    Run Connect-TestSite.ps1 first to establish a connection.
    Requires: PnP.PowerShell module with Site Owner permissions.

.EXAMPLE
    .\Connect-TestSite.ps1
    .\Setup-TestEnvironment.ps1
#>

param(
    [switch]$SkipLibraries,
    [switch]$SkipGroups,
    [switch]$SkipPermissions
)

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  External Sharing Risk Scanner - Test Setup" -ForegroundColor Cyan
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

# ============================================================
# 1. CREATE DOCUMENT LIBRARIES
# ============================================================
if (-not $SkipLibraries) {
    Write-Host "`n--- Creating Document Libraries ---" -ForegroundColor Yellow

    $libraries = @(
        @{ Title = "Contracts"; Description = "Sensitive contract documents" },
        @{ Title = "Project Files"; Description = "Project collaboration files" },
        @{ Title = "Shared Reports"; Description = "Reports shared with external partners" },
        @{ Title = "Internal Only"; Description = "Internal team documents" }
    )

    foreach ($lib in $libraries) {
        $existing = Get-PnPList -Identity $lib.Title -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [EXISTS] $($lib.Title)" -ForegroundColor Gray
        }
        else {
            New-PnPList -Title $lib.Title -Template DocumentLibrary -ErrorAction Stop | Out-Null
            Write-Host "  [CREATED] $($lib.Title)" -ForegroundColor Green
        }
    }

    # Create a custom list for tracking
    $trackingList = Get-PnPList -Identity "Risk Tracking" -ErrorAction SilentlyContinue
    if (-not $trackingList) {
        New-PnPList -Title "Risk Tracking" -Template GenericList | Out-Null
        Write-Host "  [CREATED] Risk Tracking (list)" -ForegroundColor Green
    }
    else {
        Write-Host "  [EXISTS] Risk Tracking (list)" -ForegroundColor Gray
    }

    # Upload sample documents
    $sampleContent = "This is a sample document for testing the External Sharing Risk Scanner."
    $tempFile = [System.IO.Path]::GetTempFileName() + ".txt"
    $sampleContent | Out-File -FilePath $tempFile -Encoding UTF8

    foreach ($lib in $libraries) {
        try {
            Add-PnPFile -Path $tempFile -Folder $lib.Title -NewFileName "Sample-Document.txt" -ErrorAction SilentlyContinue | Out-Null
            Write-Host "  [FILE] Added sample doc to $($lib.Title)" -ForegroundColor DarkGreen
        }
        catch {
            Write-Host "  [SKIP] Could not add file to $($lib.Title)" -ForegroundColor Gray
        }
    }

    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
}

# ============================================================
# 2. CREATE SHAREPOINT GROUPS
# ============================================================
if (-not $SkipGroups) {
    Write-Host "`n--- Creating SharePoint Groups ---" -ForegroundColor Yellow

    $groups = @(
        @{ Title = "External Partners"; Description = "Group for external partner access" },
        @{ Title = "Project Reviewers"; Description = "Mixed internal/external reviewers" },
        @{ Title = "Empty Test Group"; Description = "An empty group for testing empty detection" }
    )

    foreach ($grp in $groups) {
        $existing = Get-PnPGroup -Identity $grp.Title -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [EXISTS] $($grp.Title)" -ForegroundColor Gray
        }
        else {
            try {
                New-PnPGroup -Title $grp.Title -Description $grp.Description -ErrorAction Stop | Out-Null
                Write-Host "  [CREATED] $($grp.Title)" -ForegroundColor Green
            }
            catch {
                Write-Host "  [ERROR] Failed to create $($grp.Title): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # Note: Adding actual external users requires guest accounts in the tenant.
    # If you have external guest accounts, uncomment and modify the following:
    #
    # $externalUsers = @(
    #     "extuser1@externaldomain.com",
    #     "consultant@partnerfirm.com"
    # )
    # foreach ($extUser in $externalUsers) {
    #     try {
    #         Add-PnPGroupMember -Group "External Partners" -LoginName $extUser
    #         Write-Host "  [ADDED] $extUser to External Partners" -ForegroundColor Green
    #     } catch {
    #         Write-Host "  [ERROR] Could not add $extUser: $($_.Exception.Message)" -ForegroundColor Red
    #     }
    # }

    Write-Host "`n  NOTE: To test external user detection, manually invite external" -ForegroundColor DarkYellow
    Write-Host "  guest users to the site or add them to the 'External Partners' group." -ForegroundColor DarkYellow
}

# ============================================================
# 3. BREAK PERMISSION INHERITANCE
# ============================================================
if (-not $SkipPermissions) {
    Write-Host "`n--- Configuring Unique Permissions ---" -ForegroundColor Yellow

    # Break inheritance on "Contracts" library
    try {
        $contractsList = Get-PnPList -Identity "Contracts" -ErrorAction Stop
        Set-PnPList -Identity "Contracts" -BreakRoleInheritance -ErrorAction Stop
        Write-Host "  [DONE] Broke inheritance on 'Contracts' library" -ForegroundColor Green
    }
    catch {
        Write-Host "  [WARN] Could not break inheritance on Contracts: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Break inheritance on "Shared Reports" library
    try {
        Set-PnPList -Identity "Shared Reports" -BreakRoleInheritance -ErrorAction Stop
        Write-Host "  [DONE] Broke inheritance on 'Shared Reports' library" -ForegroundColor Green
    }
    catch {
        Write-Host "  [WARN] Could not break inheritance on Shared Reports: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # "Project Files" and "Internal Only" keep inherited permissions (for contrast)
    Write-Host "  [INFO] 'Project Files' and 'Internal Only' retain inherited permissions" -ForegroundColor Gray
}

# ============================================================
# SUMMARY
# ============================================================
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Setup Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nTest data summary:" -ForegroundColor White
Write-Host "  - 4 document libraries (2 with unique permissions)" -ForegroundColor White
Write-Host "  - 1 custom list (Risk Tracking)" -ForegroundColor White
Write-Host "  - 3 SharePoint groups (1 empty)" -ForegroundColor White
Write-Host "  - Sample documents uploaded" -ForegroundColor White
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "  1. Manually invite 1-2 external guests to the site" -ForegroundColor White
Write-Host "  2. Add external guests to 'External Partners' group" -ForegroundColor White
Write-Host "  3. Deploy the web part and test on this site" -ForegroundColor White
Write-Host "  4. Run Verify-TestData.ps1 to confirm setup" -ForegroundColor White
