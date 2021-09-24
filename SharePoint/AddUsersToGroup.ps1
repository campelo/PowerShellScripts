<#

.SYNOPSIS
Add users to a Sharepoint group.

.DESCRIPTION
Add users to a Sharepoint group.

.EXAMPLE
.\AddUsersToGroup.ps1 [FileName]
.\AddUsersToGroup.ps1 "./FileName.xlsx"

.PARAMETER FileName
Xlsx file name with all sites urls.

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 1)]
    [string] $FileName,
    [Parameter(Mandatory, Position = 2)]
    [string] $Group,
    [Parameter(Mandatory, Position = 3)]
    [string] $UserEmail
)

function AddUserToGroup {
    param(
        [Parameter(Mandatory, Position = 1)]
        [string] $GroupName,
        [Parameter(Mandatory, Position = 2)]
        [string] $EmailAddress
    )
    
    $Groups = Get-PnPGroup | Where-Object { $_.Title.Contains($GroupName) }
    ForEach ($g in $Groups) {
        Add-PnPGroupMember -EmailAddress $EmailAddress -Identity $g
    }
}

function ConnectToHost {
    param (
        [Parameter(Mandatory, Position = 1)]
        [string] $SiteUrl
    )
    try {
        #Connect to PNP Online
        Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
        Connect-PnPOnline -Interactive -Url $($SiteUrl)
    }
    catch {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
    }
}

function DisconnectHost {
    Write-Host "Disconnecting..." -ForegroundColor Cyan
    Disconnect-PnPOnline
}

function InstallPSExcel {
    try {
        #Installing PSExcel module
        Install-Module PSExcel
        Get-Command -Module PSExcel | Out-Null
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

InstallPSExcel

$objExcel = New-Excel -Path "$($FileName)"
$Worksheet = $objExcel | Get-Worksheet -Name "Metadata"
$totalNoOfRecords = $Worksheet.Dimension.Rows
$totalNoOfItems = $totalNoOfRecords - 1
$rowNo = 1
$col = 1
    
for ($i = 1; $i -le $totalNoOfItems; $i++) {
    $SiteUrl = $WorkSheet.Cells.Item($rowNo + $i, $col).text.Trim()
    if (![string]::IsNullOrEmpty($SiteUrl)) {
        try {
            ConnectToHost $SiteUrl
            AddUserToGroup "$Group" "$UserEmail"
        }
        catch {
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

DisconnectHost
