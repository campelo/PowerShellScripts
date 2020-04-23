<#

.SYNOPSIS
Change a SharePoint list's url.

.DESCRIPTION
This script will change a sharepoint list's url to avoid no-friendly chars. For example, to move my list's url from "https://MyCompany.sharepoint.com/sites/MySiteCollection/My%20List" to "https://MyCompany.sharepoint.com/sites/MySiteCollection/MyList"

.EXAMPLE
./ChangeListUrl.ps1 [SiteUrl] [ListName] [NewListUrl]

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
  
#Connect to PNP Online
Connect-PnPOnline -Url $($SiteUrl) -UseWebLogin
 
#Get the List
$List= Get-PnPList -Identity $($ListName) -Includes RootFolder
 
#sharepoint online powershell change list url
$List.Rootfolder.MoveTo($($NewListUrl))
Invoke-PnPQuery