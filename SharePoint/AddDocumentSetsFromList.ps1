<#
.SYNOPSIS
  This function will create documentsets for an existing library.

.DESCRIPTION
  This function will create documentsets for an existing library.

.PARAMETER SiteUrl
  Specifies the site address.

.PARAMETER ListName
  Specifies the list/librarie's name.

.PARAMETER DocumentSetName
  Specifies the documentset's name

.PARAMETER FileName
  Specifies file that has a list of all documentset names to create.

.EXAMPLE
  .\AddDocumentSetsFromList.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" "My Library" "DocumentSet Name" ".\FileName.txt"

.NOTES
  
#>
[CmdletBinding()]
param (
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 1)]
  [string] $SiteUrl,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 2)]
  [string] $ListName,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 3)]
  [string] $DocumentSetName,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 4)]
  [string] $FileName
)

try {
  #Connect to PNP Online
  Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
  Connect-PnPOnline -Url "$($SiteUrl)" -Interactive

  $filePath = Join-Path (Get-Location).Path $($FileName)
  $lines = [System.IO.File]::ReadLines($($filePath), [System.Text.Encoding]::Default) | Sort-Object -Property @{Expression = { $_.Trim() } } -Unique 
  foreach ($line in $lines) {
    Add-PnPDocumentSet -List "$($ListName)" -ContentType "$($DocumentSetName)" -Name "$($line)"
  }
}
catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline