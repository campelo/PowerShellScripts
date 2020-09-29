<#
.SYNOPSIS
  This function will create site columns with normalized internal names.

.DESCRIPTION
  This function will create site columns with normalized internal names.

.PARAMETER SiteUrl
  Specifies the site to add new columns

.PARAMETER FileName
  Specifies the excel file to read all new columns

.EXAMPLE
  .\CreateSiteColumnsFromXlsx.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" ".\FileName.xlsx"

.NOTES
  
#>
[CmdletBinding()]
param (
	[ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl,
	[Parameter(Mandatory = $True, Position = 2)]
    [string] $FileName
)

try {
	#Installing PSExcel module
	Install-Module PSExcel
	Get-Command -Module PSExcel
	
    #Connect to PNP Online
    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url "$($SiteUrl)" -UseWebLogin
	  
	$objExcel= New-Excel -Path "$($FileName)"
	$Worksheet = $objExcel | Get-Worksheet -Name "Metadata"
	$totalNoOfRecords= $Worksheet.Dimension.Rows
	$totalNoOfItems= $totalNoOfRecords - 1  
	$rowNo, $colType= 3, 1  
	$rowNo, $colDisplayName= 3, 2  
	if ($totalNoOfRecords -gt 1) {  
		#Loop to get values from excel file  
		for ($i= 1; $i -le $totalNoOfRecords - 1; $i++) {
			$columnDisplayName=$WorkSheet.Cells.Item($rowNo + $i, $colDisplayName).text.Trim()  
			$columnType=$WorkSheet.Cells.Item($rowNo + $i, $colType).text.Trim()
			if(![string]::IsNullOrEmpty($columnType) -and ![string]::IsNullOrEmpty($columnDisplayName)){
				$columnInternalName= & .\String-ToAlphaNumeric.ps1 -MainString "$($columnDisplayName)"
				$columnInternalName= "$($columnInternalName)".Trim()
				#Adding field
				Write-Verbose "Creating column '$($columnDisplayName)' with internal name: '$($columnInternalName)' ..."
				Add-PnPField -InternalName "$($columnInternalName)" -DisplayName "$($columnDisplayName)" -Type "$($columnType)"
			}
		}  
	}
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline