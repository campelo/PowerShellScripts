<#

.SYNOPSIS
This is a Powershell script to export a SharePoint site.

.DESCRIPTION
The script itself will export sharepoint site to a xml file.

.EXAMPLE
.\ExportSiteToFile.ps1 [SiteUrl]

.NOTES
This script will export a site with all handlers by default.
About handlers enum: https://docs.microsoft.com/fr-ca/dotnet/api/officedevpnp.core.framework.provisioning.model.handlers
About ProvisioningTemplate: https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpprovisioningtemplate

.PARAMETER SiteUrl
Url from original site collection or sub-site. Ex: "https://MyCompany.sharepoint.com/sites/MySite"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl
)

try {
    #xml(default) or pnp
    $FileName = $($SiteUrl).TrimEnd("/").Split("/")[-1]
    $OriginalSiteFileName = "$($FileName).xml"

    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url "$($SiteUrl)" -UseWebLogin

    Write-Host "Exporting site's template to '$($OriginalSiteFileName)' file..." -ForegroundColor Cyan
    #Get-PnPProvisioningTemplate -Out .\$OriginalSiteFileName -PersistBrandingFiles -ExcludeHandlers SiteSecurity, PropertyBagEntries, Navigation
    Get-PnPProvisioningTemplate -Out .\$OriginalSiteFileName -PersistBrandingFiles -Handlers All
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline
