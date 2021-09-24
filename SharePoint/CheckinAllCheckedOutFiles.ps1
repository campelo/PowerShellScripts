<#

.SYNOPSIS
Check in all checked out files.

.DESCRIPTION
This script will check in all checked out files of a library.

.EXAMPLE
.\CheckinAllCheckedOutFiles.ps1 [SiteUrl] [ListName]
.\CheckinAllCheckedOutFiles.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" "My List"

.NOTES
Source: https://www.sharepointdiary.com/2017/07/sharepoint-online-powershell-to-bulk-check-in-all-documents.html

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
    
    #Get All List Items from the List - Filter Files
    $ListItems = Get-PnPListItem -List $ListName -PageSize 500 | Where-Object {$_["FileLeafRef"] -like "*.*"}
    
    #Loop through each list item
    ForEach ($Item in $ListItems)
    {
        #Write-Verbose "Testing if file '$($Item.FieldValues["FileRef"])' is Checked-Out"
        #Get the File from List Item
        $File = Get-PnPProperty -ClientObject $Item -Property File
    
        If($File.Level -eq "Checkout")
        {
            #Check-In and Approve the File
            Set-PnPFileCheckedIn -Url $File.ServerRelativeUrl -CheckinType MajorCheckIn
    
            Write-Verbose "File Checked-In: '$($File.ServerRelativeUrl)'"
        }
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline