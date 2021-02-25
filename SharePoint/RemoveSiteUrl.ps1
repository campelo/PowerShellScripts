# https://monDomain-admin.sharepoint.com
$adminUrl = ""
# https://monDomain.sharepoint.com/sites/siteUrl
$siteUrl = ""

#Connect to service
Connect-SPOService $adminUrl

#Get site url's info.
Get-SPOSite -Identity $siteUrl
#Remove site url.
Remove-SPOSite -Identity $siteUrl