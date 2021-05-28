<#

.SYNOPSIS

.DESCRIPTION

.EXAMPLE

.NOTES
Sources: 
https://sharepointrelated.com/2011/11/28/get-sharepoint-lists-by-using-powershell/
https://www.sharepointdiary.com/2016/04/get-list-fields-in-sharepoint-using-powershell.html

.PARAMETER URL

.PARAMETER FileName

#>

[CmdletBinding()]
param
(
  [Parameter(Mandatory = $true, Position = 1)]
  [string] $URL,
  [Parameter(Position = 2)]
  [string] $FileName
)

function WriteEmptySites {
  param(
    $list,
    $fileName
  )
  $list | Sort-Object ItCount, Url | Select-Object ItCount, Url | Export-Csv -path ".\$($fileName).csv" -NoTypeInformation -Encoding Default
}

#Get all lists in farm
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

if ([string]::IsNullOrEmpty($FileName)) {
  $FileName = "EmptySites"
}
 
#variables
$siteItemsCount = 0
$sites = @()

$webs = Get-SPWeb "$($URL)/*" -Limit all -ErrorAction SilentlyContinue

if ($null -eq $webs.count -OR $webs.count -ge 1) {
  foreach ($web in $webs) {
    $lists = $web.Lists
    foreach ($list in $lists) {
      $siteItemsCount += $list.Items.count
    }
    $r = New-Object -TypeName PSObject
    $r | Add-Member -Name "Url" -MemberType Noteproperty -Value "$($web.URL)"
    $r | Add-Member -Name "ItCount" -MemberType Noteproperty -Value $siteItemsCount
    $sites += $r
    $siteItemsCount = 0
    $web.Dispose()
  }
  
  WriteEmptySites $sites $FileName
}
else {
  Write-Host "No webs retrieved, please check your permissions" -ForegroundColor Red -BackgroundColor Black
}