#Config Variables
$AdminCenterURL = "https://xxx-admin.sharepoint.com"
 
#Connect to PnP Online
#Connect-PnPOnline -Url $AdminCenterURL -Credentials (Get-Credential)
Connect-PnPOnline -Url $AdminCenterURL -Interactive
 
#Export Term Store Data to XML
Export-PnPTermGroupToXml -Out "C:\Temp\TermStoreData.xml"


#Read more: https://www.sharepointdiary.com/2016/12/sharepoint-online-powershell-to-export-term-store-data.html#ixzz6XU9YRCH6