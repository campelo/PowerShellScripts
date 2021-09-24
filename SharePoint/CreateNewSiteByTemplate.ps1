<#

.SYNOPSIS
.\CreateNewSiteByTemplate.ps1 [Url] [Titre] [CreateurID] [AdminSiteUrl]

.DESCRIPTION

.EXAMPLE
.\CreateNewSiteByTemplate.ps1 "https://mycompany.sharepoint.com/site/mysite" "My site title" "firstname.lastname@MyTenantName.onmicrosoft.com" "https://mycompany-admin.sharepoint.com"

.NOTES
Source: https://www.sharepointdiary.com/2016/10/sharepoint-online-access-request-email-settings-powershell.html

.PARAMETER Url

.PARAMETER Titre

.PARAMETER CreateurID

.PARAMETER AdminSiteUrl

#>

[CmdletBinding()]
param (
    [string] $url,
    [string] $title,
    [string] $createurID,
    [string] $adminSiteUrl,
    [string] $template, #Modern Team Site without O365 group
    [int] $timezone,
    [int] $lcid,
    [string] $templateFileName
)

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

function CreateSite {
    param (
        [Parameter(Mandatory, Position = 1)]
        [string] $SiteUrl
    )
    Return
    $url = "https://MyCompany.sharepoint.com/sites/MySite";
    $title = "My Site Title";
    $modifiedTitle = $($title -replace ",", "_")
    $ownerGroup = "Owners of $modifiedTitle"; #"Propriétaires de $modifiedTitle";
    $groupeMembre = "$modifiedTitle - Members"; #"$modifiedTitle - Membres";
    
    $createurID = "first.lastname@MyCompany.onmicrosoft.com";
    $adminSiteUrl = "https://MyCompany-admin.sharepoint.com";
    $template = "STS#3"; #Modern Team Site without O365 group
    $timezone = 10;
    $lcid = 1036;
    $templateFileName = ".\Documents.xml";
    $listToExtract = "Documents library name"
    
    #ConnectToHost $($adminSiteUrl)
    Connect-PnPOnline -Url $adminSiteUrl -Interactive
    
    #Créer une collection de sites
    New-PnPTenantSite -Url $url -Owner $createurID -Title $title -Template $template -TimeZone $timezone -Lcid $lcid
    
    #Obetnir le template
    #Get-PnPSiteTemplate -Out $templateFileName -IncludeSiteGroups -ListsToExtract $listToExtract
    
    #Appliquer un template
    #ConnectToHost $($url)
    Connect-PnPOnline -Url $url -Interactive
    Get-PnPWeb
    Invoke-PnPSiteTemplate -Path $templateFileName
    
    #Retirer les paramètres de demande d'accès
    $web = Get-PnPWeb
    # When an emailaddress is selected, empty it to disable the setting
    $web.RequestAccessEmail = ""
    # When the first radiobutton, group, is selected, setting it 
    # to false disables the setting
    $web.SetUseAccessRequestDefaultAndUpdate($false)
    $web.MembersCanShare = $False   
    $web.AssociatedMemberGroup.AllowMembersEditMembership = $False
    $web.AssociatedMemberGroup.Update()
    $web.Update()
    $web.Context.ExecuteQuery()
    
    #Retirer son compte du groupe propriétaires et Admin de la collection de sites
    Remove-PnPGroupMember -LoginName $createurID -Group $ownerGroup
    Remove-PnPSiteCollectionAdmin -Owners "i:0#.f|membership|$createurID"
    
    #Retirer le niveau d'autorisation de Modification sur le groupe membre
    Set-PnPGroupPermissions -Identity $groupeMembre -RemoveRole 'Modification'
    
    #Retirer le type de contenu par défaut afin d'utiliser uniquement ceux personnalisées
    Remove-PnPContentTypeFromList -List $listToExtract -ContentType "Document"
    #Remove-PnPList -Identity Documents -Force
    
    DisconnectHost
}

Write-Host "This script is not ready yet. Run it manually" -foregroundcolor Red