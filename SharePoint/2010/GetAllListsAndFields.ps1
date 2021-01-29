<# 

.SYNOPSIS 

.DESCRIPTION 

.EXAMPLE 

.NOTES 
Sources:  
https://sharepointrelated.com/2011/11/28/get-sharepoint-lists-by-using-powershell/ 
https://www.sharepointdiary.com/2016/04/get-list-fields-in-sharepoint-using-powershell.html 

.PARAMETER URL 

.PARAMETER FilePath 

#> 

[CmdletBinding()] 
param 
( 
  [string] $URL, 
  [string] $FilePath 
) 

function GetFieldsFromList { 
  [CmdletBinding()] 
  param ( 
    $List, 
    $ParentFolder 
  ) 

  $List.Fields | Where-Object { $_.CanBeDeleted } | Select-Object Title, Type | Export-Csv -path ".\$ParentFolder\($($List.BaseTemplate)) - $List.csv" -NoTypeInformation 
} 

function CreateFolder { 
  param($FolderName) 

  New-Item -ItemType Directory -Force -Path "$FolderName" 
} 

#Get all lists in farm 
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue 
  
#Counter variables 
$webcount = 0 
$listcount = 0 

$WriteToFile = ![string]::IsNullOrEmpty($FilePath) 

if (!$URL) { 
  #Grab all webs 
  $webs = (Get-SPSite -limit all | Get-SPWeb -Limit all -ErrorAction SilentlyContinue) 
}
else { 
  $webs = Get-SPWeb $URL 
} 

if ($webs.count -ge 1 -OR $webs.count -eq $null) { 
  foreach ($web in $webs) { 
    $newFolder = $web.Title 
    CreateFolder $newFolder 

    #Grab all lists in the current web 
    $lists = $web.Lists 
    Write-Verbose "Website $($web.url)" 
    if ($WriteToFile -eq $true) { Add-Content -Path $FilePath -Value "Website $($web.url)" } 
    foreach ($list in $lists) { 
      $listcount += 1 
      Write-Verbose " – $($list.Title)"  
      if ($WriteToFile -eq $true) { Add-Content -Path $FilePath -Value " – $($list.Title)" } 
      GetFieldsFromList $list $newFolder 
    } 
    $webcount += 1 
    $web.Dispose() 
  } 
  #Show total counter for checked webs and lists 
  Write-Verbose "Amount of webs checked: $webcount"  
  Write-Verbose "Amount of lists: $listcount" 
} 
else { 
  Write-Host "No webs retrieved, please check your permissions" -ForegroundColor Red -BackgroundColor Black 
} 