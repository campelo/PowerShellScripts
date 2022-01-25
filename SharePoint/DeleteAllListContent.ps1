<#

.SYNOPSIS
This is a Powershell script deletes all list or library's content.

.DESCRIPTION
This script will delete all list/library's content. The list/library will be empty after the script is executed.

.EXAMPLE
.\DeleteAllListContent.ps1 [SiteUrl] [ListName]

.NOTES
Source: https://vladilen.com/office-365/spo/fastest-way-to-delete-all-items-in-a-large-list/

.PARAMETER SiteUrl
Url from original site collection or sub-site. Ex: "https://MyCompany.sharepoint.com/sites/MySite"

.PARAMETER ListName
List's name. Ex: "my list"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl,
    [Parameter(Mandatory = $True, Position = 2)]
    [string] $ListName
)

try {
    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url "$($SiteUrl)" -Interactive

    Write-Host "Deleting content from '$($OriginalSiteFileName)'..." -ForegroundColor Cyan
    Get-PnPListItem -List "$($ListName)" -Fields "ID" -PageSize 100 -ScriptBlock { Param($items) $items | Sort-Object -Property Id -Descending | ForEach-Object{ $_.DeleteObject() } }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline
