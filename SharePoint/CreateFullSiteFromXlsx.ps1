<#
.SYNOPSIS
  This function will create a full site from xlsx metadata.

.DESCRIPTION
  This function will create a full site from xlsx metadata.

.PARAMETER SiteUrl
  Specifies the site to add new metadata.

.PARAMETER FileName
  Specifies the excel file to read all new metadata.

.EXAMPLE
  .\CreateFullSiteFromXlsx.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" ".\FileName.xlsx"

.NOTES
  
#>
[CmdletBinding()]
param (
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $True, Position = 1)]
	[string] $SiteUrl,
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $True, Position = 2)]
	[string] $FileName
)

try {
	& .\CreateSiteColumnsFromXlsx.ps1 -SiteUrl "$($SiteUrl)" -FileName "$($FileName)"
	& .\CreateListsFromXlsx.ps1 -SiteUrl "$($SiteUrl)" -FileName "$($FileName)"
}
catch {
	Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}