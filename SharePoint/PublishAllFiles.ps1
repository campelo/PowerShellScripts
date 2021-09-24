<#

.SYNOPSIS
Publish all draft files.

.DESCRIPTION
This script will publish all draft files of a library.

.EXAMPLE
.\PublishAllFiles.ps1 [SiteUrl] [ListName]
.\PublishAllFiles.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" "My List"

.NOTES
Source: https://www.sharepointdiary.com/2018/06/sharepoint-online-publish-document-using-powershell.html

.PARAMETER SiteUrl
Site collection full url. Ex: "https://MyCompany.sharepoint.com/sites/MySiteCollection"

.PARAMETER ListName
List's name. Ex: "My List"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl,
    [Parameter(Mandatory = $True, Position = 2)]
    [string] $ListName
)

try {
    #Connect to PnP Online
    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -Interactive

    #Get all files from the document library
    $ListItems = Get-PnPListItem -List $($ListName) -PageSize 2000 | Where-Object { $_.FileSystemObjectType -eq "File" }

    #Iterate through each file
    ForEach ($Item in $($ListItems)) {
        try {
            #Get the File from List Item
            $File = Get-PnPProperty -ClientObject $($Item) -Property File
 
            #Check if file draft
            If ($File.CheckOutType -ne "None") {
                Write-Host "Draft file: '$($File.ServerRelativeUrl)'" -BackgroundColor Red -ForegroundColor White
                Continue
            }

            #Check minor version
            If ($File.MinorVersion) {
                $File.Publish("Major version Published by Script")
                $File.Update()
                Invoke-PnPQuery
                Write-Verbose "Published file at '$($File.ServerRelativeUrl)'" 
            }
        }
        catch {
            Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline