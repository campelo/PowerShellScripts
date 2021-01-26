<#

.SYNOPSIS
This script will install PnP module for Sharepoint online

.DESCRIPTION
This script will install PnP module for Sharepoint online

.EXAMPLE
.\_installPnP.ps1

.NOTES
Probably you should enable the script execution in your computer before running this script.
Set-ExecutionPolicy -ExecutionPolicy Unrestricted

Source: https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets

#>
[CmdletBinding()]
param()

try {  
    #Install PnP module
    #Install-Module SharePointPnPPowerShellOnline
    Install-Module -Name PnP.PowerShell
    #Enable script execution (Bypass or Unrestricted)
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}