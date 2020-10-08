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
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $True, Position = 2)]
	[string] $FileName
)

$nl = [Environment]::NewLine
#To store all created fields...
$htCreatedFields = @{}
#To store all no-created fields...
$arrNoCreatedFields = @()

try {
	#Installing PSExcel module
	Install-Module PSExcel
	Get-Command -Module PSExcel | Out-Null
	
	#Connect to PNP Online
	Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
	Connect-PnPOnline -Url "$($SiteUrl)" -UseWebLogin

	$ctx = Get-PnPContext
	
	$objExcel = New-Excel -Path "$($FileName)"
	$Worksheet = $objExcel | Get-Worksheet -Name "Metadata"
	$totalNoOfRecords = $Worksheet.Dimension.Rows
	$totalNoOfColumns = $Worksheet.Dimension.Columns
	$totalNoOfItems = $totalNoOfRecords - 1  
	$rowNo = 3
	$colType = 1  
	$colDisplayName = 2

	if ($totalNoOfRecords -gt 1) {  
		$existingFields = Get-PnPField
		
		Write-Host "Creating field(s)..." -ForegroundColor Cyan
		#Getting field's name and type...  
		for ($i = 1; $i -le $totalNoOfItems; $i++) {
			$columnDisplayName = $WorkSheet.Cells.Item($rowNo + $i, $colDisplayName).text.Trim()
			$columnType = $WorkSheet.Cells.Item($rowNo + $i, $colType).text.Trim()
			if (![string]::IsNullOrEmpty($columnType) -and ![string]::IsNullOrEmpty($columnDisplayName)) {	
				$columnInternalName = & .\String-ToAlphaNumeric.ps1 -MainString "$($columnDisplayName)"
				$columnInternalName = "$($columnInternalName)".Trim()
				
				#To verify if the object already exists
				$newField = $existingFields | Where-Object { 
					($_.Internalname -eq $columnInternalName) 
					#I won't verify by existing "Display Name", because SharePoint is able to create a new field with the same existing "Display Name"
					# -or ($_.Title -eq $columnDisplayName) 
				}
				if ($NULL -ne $newField) {
					$arrNoCreatedFields += $columnDisplayName
				}
				else {
					#Adding new field
					Write-Verbose "Creating column '$($columnDisplayName)' with internal name: '$($columnInternalName)' ..."
					#Result of a field creation...
					$fResult = Add-PnPField -InternalName "$($columnInternalName)" -DisplayName "$($columnDisplayName)" -Type "$($columnType)"
					if ($NULL -eq $fResult -or $NULL -eq $fResult.Id) {
						$arrNoCreatedFields += $columnDisplayName
					}
					else {
						$htCreatedFields.Add($columnDisplayName, $fResult.Id)
					}
				}
			}
		}  
		#Are there created fields?...
		if ($htCreatedFields.Count -gt 0) {
			Write-Verbose "Created fields: $($nl) $($htCreatedFields.GetEnumerator().ForEach( { "$($_.Value) = $($_.Name) $($nl)" }))"
		}

		#Are there no-created fields?...
		if ($arrNoCreatedFields -gt 0) {
			Write-Host "No-created fields:" -BackgroundColor White -ForegroundColor Red
			Write-Host $($arrNoCreatedFields -join $nl) -BackgroundColor Red -ForegroundColor White	
		}
	}
}
catch {
	Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline