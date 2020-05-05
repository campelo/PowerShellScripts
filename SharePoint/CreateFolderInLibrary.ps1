<#

.SYNOPSIS

.DESCRIPTION

.EXAMPLE
.\CreateFolderInLibrary.ps1 [FullListUrl] [FolderName]

.NOTES
Source: https://www.sharepointdiary.com/2016/08/sharepoint-online-create-folder-using-powershell.html#ixzz6LVEbz01o

.PARAMETER FullListUrl
List or library full url. Ex: "https://MyCompany.sharepoint.com/sites/MySiteCollection/MyList"

.PARAMETER FolderName
New folder's name. Ex: "NewFolder"

#>
param (
  [Parameter(Mandatory = $True, Position = 1)]
  [string] $FullListUrl,
  [Parameter(Mandatory = $True, Position = 2)]
  [string] $FolderName
)

# [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
# [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$client = Join-path $(Get-Location) "\lib\Microsoft.SharePoint.Client.dll"
$runtime = Join-path $(Get-Location) "\lib\Microsoft.SharePoint.Client.Runtime.dll"

Import-Module $client
Import-Module $runtime
 
Try {

  $SiteURL = $($FullListUrl).TrimEnd("/").Split("/")[0..4] -Join "/"
  $ListURL = "/" + $($($FullListUrl).TrimEnd("/").Split("/")[-3..-1] -Join "/")

  Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
  Connect-PnPOnline -Url "$($SiteUrl)" -UseWebLogin

  Write-Host "Set up the context..." -ForegroundColor Cyan
  $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)

  Write-Host "Get the list root folder..." -ForegroundColor Cyan
  $ParentFolder = $Context.web.GetFolderByServerRelativeUrl($ListURL)
 
  Write-Host "Creating folder $($FolderName)..." -ForegroundColor Cyan
  $Folder = $ParentFolder.Folders.Add($FolderName)
  $ParentFolder.Context.ExecuteQuery()
 
  Write-host "New Folder Created Successfully!" -ForegroundColor Green
}
catch {
  write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline