<#
.SYNOPSIS
  This function will recursively delete a child file that has a same name of his parent folder.

.DESCRIPTION
  This function will recursively delete a child file that has a same name of his parent folder.

.PARAMETER path
  The main folder to do search for patterns

.PARAMETER fileExtName
  File extension name

.EXAMPLE
  .\DeleteChildFileSameFolderName.ps1 "c:\tmp" ".xlsx"

  This will look for child file in current directory to delete it.

.NOTES
  
#>
[CmdletBinding()]
param (
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 1)]
  [string] $path,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 2)]
  [string] $fileExtName
)

try {
  $paths = Get-ChildItem -Path "$($path)" -Recurse -Directory -Force -ErrorAction SilentlyContinue | Select-Object FullName
  
  ForEach ( $p in $paths ) {
    $curDirName = Split-Path $p.FullName -Leaf
    $sFile = "$($p.FullName)\$($curDirName)$($fileExtName)"
    if ( Test-Path -Path $sFile ) {
      Write-Verbose "Removing $($sFile)"
      Remove-Item "$($sFile)" -Force
    }
  }
}
catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}