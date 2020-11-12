<#
.SYNOPSIS
  This function will recursively move all content from a folder to another.

.DESCRIPTION
  This function will recursively move all content from a folder to another.

.PARAMETER path
  The main folder to do search for patterns

.PARAMETER targetFolder
  Destination folder to move content

.PARAMETER searchedFolderName
  Pattern of searched folder's name

.EXAMPLE
  .\MoveContentToFolder.ps1 "c:\tmp" "documents" "..\" 

  This will look in all content of tmp folder and will move all content of each documents folders to this parent's folder.

.NOTES
  
#>
[CmdletBinding()]
param (
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 1)]
  [string] $path,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 2)]
  [string] $searchedFolderName,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 3)]
  [string] $targetFolder  
)

try {
  $isRelativePath = $targetFolder.StartsWith(".")
  
  if ( !$isRelativePath ) {
    $tFolder = $targetFolder
  }

  $paths = Get-ChildItem -Path "$($path)" -Recurse -Directory -Force -ErrorAction SilentlyContinue | Select-Object FullName
  [array]::Reverse($paths)
  ForEach ( $p in $paths ) {
    if ( $p.FullName.EndsWith("$($searchedFolderName)") ) {
      if ($isRelativePath){
        $tFolder = "$($p.FullName)\$($targetFolder)"
      }
      Move-Item "$($p.FullName)\*" "$($tFolder)"
    }
  }
}
catch {
  Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}