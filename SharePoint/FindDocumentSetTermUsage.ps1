<#

.SYNOPSIS
Find all term usages.

.DESCRIPTION
This script will list all sharepoint sites that are using a document-set term.

.EXAMPLE
.\FindDocumentSetTermUsage.ps1 [SiteUrl] [TermId]

.NOTES
Source: https://www.techmikael.com/2018/05/locating-where-term-set-is-used-in.html

.PARAMETER SiteUrl
Site collection full url. Ex: "https://MyCompany.sharepoint.com/sites/MySiteCollection"

.PARAMETER TermId
Term's id. Ex: "09c963d1-662e-4fb9-a8e3-58b52dc7bd83"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl,
    [Parameter(Mandatory = $True, Position = 2)]
    [string] $TermId
)

try {  
    #Connect to PNP Online
    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $($SiteUrl) -Interactive

    #Submiting query
    Write-Host "Finding all term usages... '$($TermId)'..." -ForegroundColor Cyan
    Submit-PnPSearchQuery -Query "$($TermId)" -CollapseSpecification "SPSiteUrl:1" -RelevantResults -SelectProperties "SPSiteUrl" | Select-Object SPSiteUrl
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline