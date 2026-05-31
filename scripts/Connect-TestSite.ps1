<#
.SYNOPSIS
    Connects to the External Sharing Risk Scanner test site using PnP PowerShell.

.DESCRIPTION
    Establishes an interactive connection to the SharePoint Online test site
    using the registered Azure AD app for PnP PowerShell.

.NOTES
    Requires: PnP.PowerShell module
    Install: Install-Module PnP.PowerShell -Scope CurrentUser
#>

$SiteURL = "https://7ktf8h.sharepoint.com/sites/external-sharing-risk-lab"
$ClientId = "e4eef86c-f84c-4382-bfd0-935b619fc06c"

Write-Host "Connecting to: $SiteURL" -ForegroundColor Cyan
Write-Host "Using Client ID: $ClientId" -ForegroundColor Gray

Connect-PnPOnline -Url $SiteURL -Interactive -ClientId $ClientId

if (Get-PnPContext) {
    Write-Host "✓ Connected successfully!" -ForegroundColor Green
    Write-Host "Site: $(Get-PnPWeb | Select-Object -ExpandProperty Title)" -ForegroundColor Green
}
else {
    Write-Host "✗ Connection failed." -ForegroundColor Red
}
