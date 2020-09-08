#Config Variables
$AdminCenterURL = "https://xxx-admin.sharepoint.com"
$TermSetId="404cfd9f-8c74-4fbd-86d7-8ba39e66ddd7"
$FilePath="C:\Temp\TermsetData.txt"
 
#Connect to PnP Online
#Connect-PnPOnline -Url $AdminCenterURL -Credentials (Get-Credential)
Connect-PnPOnline -Url $AdminCenterURL -UseWebLogin
 
#Export Term set
Export-PnPTaxonomy -TermSetId $TermsetID -Path $FilePath


#Read more: https://www.sharepointdiary.com/2016/12/sharepoint-online-powershell-to-export-term-set-to-csv.html#ixzz6XUEYXPEf