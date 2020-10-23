<#
.SYNOPSIS
  This function will rename all folders from a from-to excel's table.

.DESCRIPTION
  This function will rename all folders from a from-to excel's table. The excel's table should has the current name on the first column and new name on the second.

.PARAMETER FileName
  Specifies the excel file to read all from-to relation.

.EXAMPLE
  .\RenameFoldersFromXlsx.ps1 ".\FileName.xlsx"

.NOTES
  
#>
[CmdletBinding()]
param (
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $True, Position = 1)]
	[string] $FileName
)

try {
	#Installing PSExcel module
	Install-Module PSExcel
	Get-Command -Module PSExcel | Out-Null
	
	$objExcel = New-Excel -Path "$($FileName)"
	$Worksheet = $objExcel | Get-Worksheet
	$totalNoOfRecords = $Worksheet.Dimension.Rows
	$rowNo = 1
	$colOldName = 1
	$colNewName = 2

	if ($totalNoOfRecords -gt 1) {  
		for ($i = 0; $i -lt $totalNoOfRecords; $i++) {

			#Getting old names...
			$oldName = $WorkSheet.Cells.Item($rowNo + $i, $colOldName).text.Trim()
			$newName = $WorkSheet.Cells.Item($rowNo + $i, $colNewName).text.Trim()

			if (!([string]::IsNullOrEmpty($oldName) -Or [string]::IsNullOrEmpty($newName))) {	
				try {
					Write-Verbose "Renaming folder '$($oldname)' to '$($newName)'"
					Rename-Item ".\$($oldName)" "$($newName)"
				}
				catch {
					Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
				}
			}
		}
	}
}
catch {
	Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}