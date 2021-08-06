<#

.SYNOPSIS

.DESCRIPTION

.EXAMPLE

.NOTES
Sources: 
https://sharepointrelated.com/2011/11/28/get-sharepoint-lists-by-using-powershell/
https://www.sharepointdiary.com/2016/04/get-list-fields-in-sharepoint-using-powershell.html

.PARAMETER URL

#>

[CmdletBinding()]
param
(
  [Parameter(Mandatory = $true, Position = 1)]
  [string] $URL
)

#Get all lists in farm
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
 
#Counter variables
$webs = Get-SPWeb $URL

if ($null -eq $webs.count -OR $webs.count -ge 1) {
  foreach ($web in $webs) {
    $fileName = "Libraries from $($web.Title).txt" 
	
    #Grab all lists in the current web
    $libraries = $web.Lists | Where-Object { $_.BaseType -eq "DocumentLibrary" }
    Write-Verbose "Website $($web.url)"
    foreach ($lib in $libraries) {
      $lib.Title >> "$($fileName)"
    }
    $web.Dispose()
  }
}
else {
  Write-Host "No webs retrieved, please check your permissions" -ForegroundColor Red -BackgroundColor Black
}