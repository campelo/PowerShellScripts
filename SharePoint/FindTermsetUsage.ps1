<#

.SYNOPSIS
Find sites are using a specific termset.

.DESCRIPTION
This script will search all sharepoint sites that is using a specific termset.

.EXAMPLE
.\FindTermsetUsage.ps1 [SiteUrl] [TermsetGuid]
.\FindTermsetUsage.ps1 "https://MyCompany.sharepoint.com/" "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"

.NOTES
Source: https://www.techmikael.com/2018/05/locating-where-term-set-is-used-in.html

.PARAMETER SiteUrl
Sharepoint tenant's url. Ex: "https://MyCompany.sharepoint.com/"

.PARAMETER TermsetGuid
Termset GUID. Ex: "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl,
    [Parameter(Mandatory = $True, Position = 2)]
    [string] $TermsetGuid
)

try {  
    #Connect to PNP Online
    Write-Host "Connecting to sharepoint '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $($SiteUrl) -UseWebLogin
 
    #Get the List
    Write-Host "Searching usages for termset..." -ForegroundColor Cyan
    Submit-PnPSearchQuery -Query "$($TermsetGuid)" -CollapseSpecification "SPSiteUrl:1" -RelevantResults -SelectProperties "SPSiteUrl" | Select-Object SPSiteUrl
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline