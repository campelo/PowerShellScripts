<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER SiteUrl

.PARAMETER FileName

.EXAMPLE

.NOTES
  
#>
[CmdletBinding()]
param (
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $True, Position = 1)]
	[string] $SiteUrl,
	[ValidateNotNullOrEmpty()]
	[Parameter(Mandatory = $True, Position = 2)]
	[string] $GroupName
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

ConnectToHost $SiteUrl

$fields = Get-PnPField -Group $GroupName | Where-Object { $_.CanBeDeleted } | Select-Object InternalName, Title, TypeDisplayName
ForEach($f in $fields) {
	$delete = Read-Host "Delete this field?" $f
	if ($delete -eq "y"){
		Remove-PnPField -Identity $f.InternalName
	}
}

DisconnectHost