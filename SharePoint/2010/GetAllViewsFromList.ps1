$web = Get-SPWeb "https://portal/sites/siteName"
foreach ($list in $web.Lists)
{
foreach ($view in $list.Views)
{
	$spView = $web.GetViewFromUrl($view.Url)
	Write-Host $spView
	foreach ($spField in $spView.ViewFields)
	{
		Write-Host "  -$($spField)"
	}
}
}