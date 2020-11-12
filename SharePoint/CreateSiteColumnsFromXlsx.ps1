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
	
	$objExcel = New-Excel -Path "$($FileName)"
	$Worksheet = $objExcel | Get-Worksheet -Name "Metadata"
	$totalNoOfRecords = $Worksheet.Dimension.Rows
	$totalNoOfItems = $totalNoOfRecords - 1  
	$rowNo = 3
	$colType = 1  
	$colDisplayName = 2

	if ($totalNoOfRecords -gt 1) {  
		
		Write-Host "Creating field(s)..." -ForegroundColor Cyan
		#Getting field's name and type...  
		for ($i = 1; $i -le $totalNoOfItems; $i++) {
			$columnDisplayName = $WorkSheet.Cells.Item($rowNo + $i, $colDisplayName).text.Trim()
			$columnType = $WorkSheet.Cells.Item($rowNo + $i, $colType).text.Trim()
			if (![string]::IsNullOrEmpty($columnType) -and ![string]::IsNullOrEmpty($columnDisplayName)) {	
				$columnInternalName = & .\String-ToAlphaNumeric.ps1 -MainString "$($columnDisplayName)"
				$columnInternalName = "$($columnInternalName)".Trim()
				
				$newField = $nIndex = 0
				$tmpName = "$($columnInternalName)"

				While ($NULL -ne $newField) {
					$newField = Get-PnPField | Where-Object { 
						($_.Internalname -eq "$($tmpName)")
						#I won't verify by existing "Display Name", because SharePoint is able to create a new field with the same existing "Display Name"
						# -or ($_.Title -eq $columnDisplayName) 
					}
					Write-Host "Field: $($newField)"
					If ($NULL -ne $newField) {
						$nIndex += 1
						$tmpName = "$($columnInternalName)$($nIndex)"
					}
				}
				#To verify if the object already exists
				
				#Adding new field
				Write-Verbose "Creating column '$($columnDisplayName)' with internal name: '$($tmpName)' ..."
				
				#Result of a field creation...
				$fResult = Add-PnPField -InternalName "$($tmpName)" -DisplayName "$($columnDisplayName)" -Type "$($columnType)"
				if ($NULL -eq $fResult -or $NULL -eq $fResult.Id) {
					$arrNoCreatedFields += $columnDisplayName
				}
				else {
					$htCreatedFields.Add($columnDisplayName, $fResult.Id)
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