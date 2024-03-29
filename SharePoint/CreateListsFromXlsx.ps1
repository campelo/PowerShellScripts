<#
.SYNOPSIS
  This function will create lists and libraries with normalized internal names.

.DESCRIPTION
  This function will create lists and libraries with normalized internal names.

.PARAMETER SiteUrl
  Specifies the site to include new lists/libraries

.PARAMETER FileName
  Specifies the excel file to read all new lists/libraries

.EXAMPLE
  .\CreateListsFromXlsx.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" ".\FileName.xlsx"

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

$htFieldsToAdd = @{}

try {
	#Installing PSExcel module
	Install-Module PSExcel
	Get-Command -Module PSExcel | Out-Null
	
	#Connect to PNP Online
	Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
	Connect-PnPOnline -Url "$($SiteUrl)" -Interactive

	$ctx = Get-PnPContext
	
	$objExcel = New-Excel -Path "$($FileName)"
	$Worksheet = $objExcel | Get-Worksheet -Name "Metadata"
	$totalNoOfRecords = $Worksheet.Dimension.Rows
	$totalNoOfColumns = $Worksheet.Dimension.Columns
	$totalNoOfLists = $totalNoOfColumns - 6
	$totalNoOfItems = $totalNoOfRecords - 1  
	$rowNo = 3
	$colDisplayName = 2
	$rowList = 2
	$colList = 6

	if ($totalNoOfRecords -gt 1) {  
		for ($($i = 0; $listTemplate = "DocumentLibrary"; $listOrLib = "library(ies)"; $vName = "Tous les documents"; $firstColView = "Nom"); 
			$i -lt 2; 
			$($i++; $listTemplate = "GenericList"; $listOrLib = "list(s)"; $vName = ""; $firstColView = "Titre")) {
			
			Write-Host "Creating $($listOrLib)..." -ForegroundColor Cyan

			#Getting libraries and lists...
			for ($col = 1; $col -le $totalNoOfLists; $col++) {
				$listDisplayName = $WorkSheet.Cells.Item($rowList + $i, $colList + $col).text.Trim()

				#Should we create list/library?...
				if (![string]::IsNullOrEmpty($listDisplayName)) {	
					$listInternalName = & .\String-ToAlphaNumeric.ps1 -MainString "$($listDisplayName)"
					$listInternalName = "$($listInternalName)".Trim()
					if ($listInternalName.Length -gt 50) {
						$listInternalName = $listInternalName.Substring(0, 50)
					}

					#Creating list...
					Write-Verbose "Creating $($listOrLib) $($listDisplayName)"
					New-PnPList -Title "$($listDisplayName)" -Url "lists/$($listInternalName)" -Template "$($listTemplate)" -OnQuickLaunch
					
					$l = Get-PnPList -Identity "lists/$($listInternalName)"

					for ($j = 1; $j -le $totalNoOfItems; $j++) {
						$fIndex = $WorkSheet.Cells.Item($rowNo + $j, $colList + $col).text.Trim()
						if (![string]::IsNullOrEmpty($fIndex)) {
							$columnDisplayName = $WorkSheet.Cells.Item($rowNo + $j, $colDisplayName).text.Trim()
							$columnInternalName = & .\String-ToAlphaNumeric.ps1 -MainString "$($columnDisplayName)"
							$columnInternalName = "$($columnInternalName)".Trim()
							$f = Get-PnPField | Where-Object { $_.Title -eq "$($columnDisplayName)" -and $_.InternalName.StartsWith("$($columnInternalName)") -and $_.Group.Contains("personnalis") } | Sort-Object -Property InternalName -Descending | Select-Object -First 1
							if ($NULL -eq $f) {
								Write-Host "Column $($columnInternalName) not found. Trying to use '$($columnDisplayName)' column instead." -ForegroundColor Yellow
								$f = Get-PnPField -Identity "$($columnDisplayName)"
							}
							$htFieldsToAdd.Add($fIndex -as [int], $f)
						}
					}
					
					$viewColumns = @()
					if ($listTemplate -eq "DocumentLibrary") {
						$viewColumns += "Type"
					}
					$viewColumns += $firstColView

					$($htFieldsToAdd.GetEnumerator() | Sort-Object -Property key).ForEach( { 
							$l.Fields.Add($_.Value) | Out-Null
							$viewColumns += "$($_.Value.InternalName)"
						})
					$l.Update()
					#$bugFix to suppress those message errors. 
					#Sources
					#https://github.com/pnp/PnP-PowerShell/issues/722
					#https://techcommunity.microsoft.com/t5/sharepoint-developer/add-spofile-fails-with-format-default-the-collection-has-not/m-p/14323
					if (![string]::IsNullOrEmpty($vName)) {
						$bugFix = Set-PnPView -List "lists/$($listInternalName)" -Identity "$($vName)" -Fields $viewColumns
					}
					else {
						$lAux = Get-PnPView -List "lists/$($listInternalName)"
						$bugFix = Set-PnPView -List "lists/$($listInternalName)" -Identity $lAux.Id -Fields $viewColumns
					}
					$htFieldsToAdd.Clear() | Out-Null
				}
			}
		}
		$ctx.ExecuteQuery()
	}
}
catch {
	Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline