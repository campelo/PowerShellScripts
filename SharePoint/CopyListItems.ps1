<#

.SYNOPSIS
Copy list items from a site to another.

.DESCRIPTION
Copy list items from a site to another. Simple content will be copied. This script was not made to complex content migration.

.EXAMPLE
.\ChangeListUrl.ps1 [Site1] [List1] [Site2] [List2]
.\ChangeListUrl.ps1 "https://MyCompany.sharepoint.com/sites/SourceSiteCollection" "Source List" "https://MyCompany.sharepoint.com/sites/TargetSiteCollection" "Target List"

.PARAMETER Site1
Source site

.PARAMETER List1
Source list

.PARAMETER Site2
Target site

.PARAMETER List2
Target list. No required if target list has the same name that source list.

#>

[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 1)]
    [string] $Site1,
    [Parameter(Mandatory, Position = 2)]
    [string] $List1,
    [Parameter(Mandatory, Position = 3)]
    [string] $Site2,
    [Parameter(Position = 4)]
    [string] $List2
)

try {  
    if([string]::IsNullOrEmpty($List2)){
        $List2 = $List1
    }

    #Connect to PNP Online
    Write-Host "Connecting to site '$($Site1)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Interactive -Url $($Site1)
    $fields = Get-PnPField -List $($List1) | Where-Object { $_.CanBeDeleted }
    $items = Get-PnPListItem -List $($List1)

    #Connect to PNP Online
    Write-Host "Connecting to site '$($Site2)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Interactive -Url $($Site2)
    foreach ($it in $items) {
        $value = @{}
        $value.Add("Title", $it['Title'])
        foreach ($f in $fields) {
            $k = $f.InternalName
            $v = $it[$k]
            if (![string]::IsNullOrEmpty($v) -and $v.GetType().ToString().StartsWith("System")) {
                $value.Add($k, $v)   
            } 
        }
        Add-PnPListItem -List $($List2) -Values $value
    }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline
