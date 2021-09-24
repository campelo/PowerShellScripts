<#

.SYNOPSIS
This is a Powershell script to export a SharePoint termsets.

.DESCRIPTION
The script itself will export SharePoint termsets to a csv file (Termsets.csv default file name).

.EXAMPLE
.\ExportTermsSet.ps1 [SiteUrl] [TermGroup] [TermSet]

.NOTES
This script will export a site terms set.

.PARAMETER SiteUrl
Url from original site collection or sub-site. Ex: "https://MyCompany.sharepoint.com/sites/MySite"

.PARAMETER TermGroup
Term group name. Ex: "Country"

.PARAMETER TermSet
Termset name. Ex: "USA"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl,
    [Parameter(Mandatory = $True, Position = 2)]
    [string] $TermGroup,
    [Parameter(Mandatory = $True, Position = 3)]
    [string] $TermSet
)

try {
    [System.Collections.ArrayList] $completeTerm = [System.Collections.ArrayList]::new()

    #File name
    $FileName = "Termsets.csv" #"$($TermGroup)_$($TermSet).csv"

    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url "$($SiteUrl)" -Interactive

    Write-Host "Finding site's terms set..." -ForegroundColor Cyan
    $termStores = Get-PnPTerm -TermGroup $($TermGroup) -TermSet $($TermSet) -IncludeChildTerms 
    
    $completeTerm = @()

    Write-Host "Searching all terms..." -ForegroundColor Cyan
    foreach ($termStore in $($termStores)) {
        $completeTerm += New-Object PsObject -Property @{
            "Term Set Name" = $($TermSet);
            "Level 1 Term"  = $($termStore).Name;
        }
        if ($($termStore).TermsCount -gt 0) {
            foreach ($termStoreChild in $($termStore).Terms) {
                $completeTerm += New-Object PsObject -Property @{ 
                    "Term Set Name" = $($TermSet);
                    "Level 1 Term"  = $($termStore).Name;
                    "Level 2 Term"  = $($termStoreChild).Name;
                }	
            }
        }
    }
    
    Write-Host "Saving file $($FileName)..." -ForegroundColor Cyan
    $completeTerm | Select-Object "Term Set Name", "Term Set Description", "LCID", "Available for Tagging", "Term Description", "Level 1 Term", "Level 2 Term", "Level 3 Term", "Level 4 Term" | Export-Csv ".\$($FileName)" -NoTypeInformation
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline