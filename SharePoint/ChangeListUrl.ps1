<#

.SYNOPSIS
Change a SharePoint list's url.

.DESCRIPTION
This script will change a sharepoint list's url to avoid no-friendly chars. For example, to move my list's url from "https://MyCompany.sharepoint.com/sites/MySiteCollection/My%20List" to "https://MyCompany.sharepoint.com/sites/MySiteCollection/MyList"

.EXAMPLE
.\ChangeListUrl.ps1 [SiteUrl] [ListName] [NewListUrl]

.NOTES
Source: https://www.sharepointdiary.com/2017/09/sharepoint-online-change-list-document-library-url-using-powershell.html#ixzz6KRtYz1tm

.PARAMETER SiteUrl
Site collection full url. Ex: "https://MyCompany.sharepoint.com/sites/MySiteCollection"

.PARAMETER ListName
List's name. Ex: "My List"

.PARAMETER NewListUrl
New list's url. Ex: "MyList"

#>
param (
    [Parameter(Mandatory=$True, Position=1)]
    [string] $SiteUrl,
    [Parameter(Mandatory=$True, Position=2)]
    [string] $ListName,
    [Parameter(Mandatory=$True, Position=3)]
    [string] $NewListUrl
)

try{  
    #Connect to PNP Online
    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $($SiteUrl) -UseWebLogin
 
    #Get the List
    Write-Host "Getting identity from list/library '$($ListName)'..." -ForegroundColor Cyan
    $List= Get-PnPList -Identity $($ListName) -Includes RootFolder
 
    #sharepoint online powershell change list url
    Write-Host "Changing list/library url to '$($NewListUrl)'..." -ForegroundColor Cyan
    $List.Rootfolder.MoveTo($($NewListUrl))
    Invoke-PnPQuery
}
catch{
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline