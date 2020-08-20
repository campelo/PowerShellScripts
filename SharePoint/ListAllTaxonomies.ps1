param (
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $SiteUrl
)

try {
    # foreach ($item in $sitesdata) {
    Write-Host "Started Traversing Site " $($SiteUrl) -ForegroundColor Yellow
    Connect-PnPOnline -URL $($SiteUrl) -useWebLogin
    $site = Get-PnPSite
    if ($site.Url -eq $($SiteUrl)) {   
        $lists = Get-PnPList
        Write-Host "Found " $lists.Count " Lists. Processing.. Please wait.."
        foreach ($list in $lists) {
            $taxonomyfields = Get-PnPField -List $list | Where-Object { $_.TypeAsString -eq "TaxonomyFieldType" }
            if ($taxonomyfields.Count -gt 0) {
                foreach ($field in $taxonomyfields) {
                    $xml = [XML]$field.SchemaXml
                    $TermSetId = ($xml | Select-Xml "//Name[text()='TermSetId']/following-sibling::Value/text()").Node.Value
                    $TermId = ($xml | Select-Xml "//Name[text()='AnchorId']/following-sibling::Value/text()").Node.Value
                    $filterData = $taxonomydata | Where-Object { $_.TermsetID -eq $TermSetId -and (($_.Level1TermId -eq $TermId) -or ($_.Level2TermId -eq $TermId) -or ($_.Level3TermId -eq $TermId)) }
                    $GroupId = $filterData[0].GroupID;
                    $GroupName = $filterData[0].Group;
                    $TermsetName = $filterData[0].TermSet;
                    $Level1TermName = $filterData[0].Level1Term;
                    $obj = New-Object PSObject;
                    $obj | Add-Member NoteProperty  ID  $($count);
                    $obj | Add-Member NoteProperty  SiteUrl  $($SiteUrl);
                    $obj | Add-Member NoteProperty  TaxonomyPresence "Present";
                    $obj | Add-Member NoteProperty  ListName  $($list.Title);
                    $obj | Add-Member NoteProperty  FieldName  $($field.Title);
                    $obj | Add-Member NoteProperty  FieldInternalName  $($field.InternalName);
                    $obj | Add-Member NoteProperty  GroupId  $($GroupId);
                    $obj | Add-Member NoteProperty  GroupName  $($GroupName);
                    $obj | Add-Member NoteProperty  TermSetId  $($TermSetId);
                    $obj | Add-Member NoteProperty  TermSetName  $($TermsetName);
                    $obj | Add-Member NoteProperty  TermId  $($TermId);
                    $obj | Add-Member NoteProperty  TermName  $($Level1TermName);
                    $results += $obj;
                }
            }
            else {
                $obj = New-Object PSObject;
                $obj | Add-Member NoteProperty  ID  $($count);
                $obj | Add-Member NoteProperty  SiteUrl  $($SiteUrl);
                $obj | Add-Member NoteProperty  TaxonomyPresence "No Taxonomy fields found";
                $obj | Add-Member NoteProperty  ListName  $($list.Title);
                $obj | Add-Member NoteProperty  FieldName  "";
                $obj | Add-Member NoteProperty  FieldInternalName  "";
                $obj | Add-Member NoteProperty  GroupId  "";
                $obj | Add-Member NoteProperty  GroupName  "";
                $obj | Add-Member NoteProperty  TermSetId  "";
                $obj | Add-Member NoteProperty  TermSetName  "";
                $obj | Add-Member NoteProperty  TermId  "";
                $obj | Add-Member NoteProperty  TermName  "";
                $results += $obj;
            }
        }
    }
    else {
        $obj = New-Object PSObject;
        $obj | Add-Member NoteProperty  ID  $($count);
        $obj | Add-Member NoteProperty  SiteUrl  $($SiteUrl);
        $obj | Add-Member NoteProperty  TaxonomyPresence "Site is not accessible";
        $obj | Add-Member NoteProperty  ListName  "";
        $obj | Add-Member NoteProperty  FieldName  "";
        $obj | Add-Member NoteProperty  FieldInternalName  "";
        $obj | Add-Member NoteProperty  GroupId  "";
        $obj | Add-Member NoteProperty  GroupName  "";
        $obj | Add-Member NoteProperty  TermSetId  "";
        $obj | Add-Member NoteProperty  TermSetName  "";
        $obj | Add-Member NoteProperty  TermId  "";
        $obj | Add-Member NoteProperty  TermName  "";
        $results += $obj;
    }
    $count++;
    Write-Host "Processing Finished .."
    # }

    if ($results.Count -gt 0) {
        Write-Host "Creating Report" -ForegroundColor Red
        $fileName = "SiteCollectionTaxonomyFields.csv"
        Write-Host "Exporting file " $fileName
        $results | Export-Csv $fileName -NoTypeInformation
        Write-Host "Creating Report Completed" -ForegroundColor Green
    }
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}

Write-Host "Disconnecting..." -ForegroundColor Cyan
Disconnect-PnPOnline