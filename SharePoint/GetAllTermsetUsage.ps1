<#

.SYNOPSIS
Get all termset usage from all sites.

.DESCRIPTION
This script will search all sharepoint sites that is using termsets.

.EXAMPLE
.\GetAllTermsetUsage.ps1 [SiteUrl]
.\GetAllTermsetUsage.ps1 "https://MyCompany.sharepoint.com/"

.NOTES
Source: https://www.techmikael.com/2018/05/locating-where-term-set-is-used-in.html

.PARAMETER SiteUrl
Sharepoint tenant's url. Ex: "https://MyCompany.sharepoint.com"

#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl
)

try {  
    #Connect to PNP Online
    Write-Host "Connecting to sharepoint '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $($SiteUrl) -Interactive

    $sites = Get-PnPTenantSite | Where-Object { $_.Template -eq "STS#3" -or $_.Template -eq "SITEPAGEPUBLISHING#0" }

    foreach ($site in $sites) {
        Connect-PnPOnline -Url "$($site.Url)" -Interactive
        $lists = Get-PnPList
        foreach ($list in $lists) {
            $taxonomyfields = Get-PnPField -List $list | Where-Object { $_.TypeAsString -eq "TaxonomyFieldType" }
            if ($taxonomyfields.Count -gt 0) {
                foreach ($field in $taxonomyfields) {
                    $xml = [XML]$field.SchemaXml
                    $TermSetId = ($xml | Select-Xml "//Name[text()='TermSetId']/following-sibling::Value/text()").Node.Value
                    "TermSetId: $($TermSetId)  => SiteUrl: $($site.Url)  => $($list.Title)"  >> .\allTermsSets.txt
                }
            }
        }
    }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline