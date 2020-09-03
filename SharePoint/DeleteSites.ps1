<#

.SYNOPSIS
This is a Powershell script to delete a list of websites.

.DESCRIPTION
This script will delete all listed websites in a file.

.EXAMPLE
.\DeleteSites.ps1 [SitesFileName]

.NOTES
This script will delete all listed websites in a file.
About Remove-SPWeb: https://docs.microsoft.com/en-us/powershell/module/sharepoint-server/remove-spweb?view=sharepoint-ps

.PARAMETER SitesFileName
A relative or absolute file name containing all websites' urls. Ex: "c:\folder\urls.txt" or "..\folder2\urls.txt"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $False, Position = 1)]
    [string] $SitesFileName = ".\list.txt"
)

try {   
    Write-Host "Reading all websites from file '$($FileName)'..." -ForegroundColor Cyan
    Get-Content $($FileName) | ForEach-Object {
		try {
			Write-Host "Removing '$($FileName)'..." -ForegroundColor Cyan
			Remove-SPWeb -Identity $($_) -Recycle:$True -Confirm:$False
		}
		catch {
			Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
		}
	}
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -foregroundcolor Red
}
