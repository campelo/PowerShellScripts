<#

.SYNOPSIS
Create a new group for all sites in a xlsx file.

.DESCRIPTION
Create a new group for all sites in a xlsx file.

.EXAMPLE
.\AddGroupGestionAcces.ps1 [FileName]
.\AddGroupGestionAcces.ps1 "./FileName.xlsx"

.PARAMETER FileName
Xlsx file name with all sites urls.

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 1)]
    [string] $FileName
)

function AddGroupToSite {
    param(
        [Parameter(Mandatory, Position = 1)]
        [string] $SiteUrl,
        [Parameter(Mandatory, Position = 2)]
        [string] $NewGroupName
    )

    ConnectToHost $SiteUrl
    
    DisableAccessRequest

    $NewGroup = New-PnPGroup -Title $NewGroupName
    Set-PnPGroup -Identity $NewGroup -AddRole "Collaboration" -Owner $NewGroup.Title -AllowMembersEditMembership $False
    
    $MemberGroup = Get-PnPGroup | Where-Object { $_.Title.Contains("Membres") }
    ForEach ($g in $MemberGroup) {
        Set-PnPGroup -Identity $g -AddRole "Collaboration" -RemoveRole "Modification" -Owner $NewGroup.Title -AllowMembersEditMembership $False
    }
    $VisitorGroup = Get-PnPGroup | Where-Object { $_.Title.Contains("Visiteurs") }
    ForEach ($g in $VisitorGroup) {
        Set-PnPGroup -Identity $g -Owner $NewGroup.Title -AllowMembersEditMembership $False
    }
}

function DisableAccessRequest {
    try {
        #Get the Web
        $Web = Get-PnPWeb
 
        #Disable access requests
        $Web.RequestAccessEmail = ""
        $Web.SetUseAccessRequestDefaultAndUpdate($False)
        #Disable members to share site and individual files and subfolders
        $Web.MembersCanShare = $False
        $Web.Update()
        $Web.Context.ExecuteQuery()    
    }
    catch {
        write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
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
        Connect-PnPOnline -UseWebLogin -Url $($SiteUrl)
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
    $siteUrl = $WorkSheet.Cells.Item($rowNo + $i, $col).text.Trim()
    if (![string]::IsNullOrEmpty($siteUrl)) {
        try {
            AddGroupToSite $siteUrl "Gestion accès"
        }
        catch {
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
        DisconnectHost
    }
}
