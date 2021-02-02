<#

.SYNOPSIS

.DESCRIPTION

.EXAMPLE

.PARAMETER FileName

.PARAMETER GroupGUID

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 1)]
    [string] $SiteUrl,
    [Parameter(Mandatory, Position = 2)]
    [string] $TermGroupGUID,
    [Parameter(Mandatory, Position = 3)]
    [string] $TermSetGUID,
    [Parameter(Mandatory, Position = 4)]
    [string] $FileName
)

function AddNewTerm {
    param(
        [Parameter(Mandatory, Position = 1)]
        [string] $TermGroupGUID,
        [Parameter(Mandatory, Position = 2)]
        [string] $TermSetGUID,
        [Parameter(Mandatory, Position = 3)]
        [string] $NewTermName
    )
    try {
        New-PnPTerm -TermGroup "$($TermGroupGUID)" -TermSet "$($TermSetGUID)" -Name "$($NewTermName)"
        Write-Verbose "Adding new term : $NewTermName"
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
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

$filePath = Join-Path (Get-Location).Path $($FileName)
   
$lines = [System.IO.File]::ReadLines($($filePath), [System.Text.Encoding]::Default)

ConnectToHost $($SiteUrl)

foreach ($line in $($lines)) {
    AddNewTerm "$($TermGroupGUID)" "$($TermSetGUID)" "$($line)"
}

DisconnectHost