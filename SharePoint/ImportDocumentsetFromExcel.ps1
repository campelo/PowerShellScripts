<#
.SYNOPSIS
  This function will create documentsets for an existing library.

.DESCRIPTION
  This function will create documentsets for an existing library.

.PARAMETER SiteUrl
  Specifies the site address.

.PARAMETER ListName
  Specifies the list/librarie's name.

.PARAMETER DocumentSetName
  Specifies the documentset's name

.PARAMETER FilesFolder
  The folder that has all files to import data.

.PARAMETER FilesFilter
  The file's extension.

.PARAMETER CreateOnly
  This parameter specifies that this script will only create Documentsets.

.PARAMETER UpdateOnly
  This parameter specifies that this script will only update Documentsets.

.PARAMETER CheckOnly
  Check all files. It doesn't make any import.

.EXAMPLE
  .\ImportDocumentsetFromExcel.ps1 "https://MyCompany.sharepoint.com/sites/MySiteCollection" "My Library" "DocumentSet Name" ".\MyFolder" "*.xlsx" -CreateOnly -UpdateOnly -CheckOnly

.NOTES
  
#>
[CmdletBinding()]
param (
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 1)]
  [string] $SiteUrl,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 2)]
  [string] $ListName,
  [ValidateNotNullOrEmpty()]
  [Parameter(Mandatory = $True, Position = 3)]
  [string] $DocumentSetName,
  [Parameter(Position = 4)]
  [string] $FilesFolder = ".\",
  [Parameter(Position = 5)]
  [string] $FilesFilter = "*.xlsx",
  [switch] $CreateOnly,
  [switch] $UpdateOnly,
  [switch] $CheckOnly
)

function ConnectToHost {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, Position = 1)]
    [string] $SiteUrl
  )
  try {
    #Connect to PNP Online
    Write-Host "Connecting to site '$($SiteUrl)'..." -ForegroundColor Cyan
    Connect-PnPOnline -UseWebLogin -Url $($SiteUrl)
  }
  catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
  }
}

function DisconnectHost {
  Write-Host "Disconnecting..." -ForegroundColor Cyan
  Disconnect-PnPOnline
}

function InstallPSExcel {
  #Installing PSExcel module
  Install-Module PSExcel
  Get-Command -Module PSExcel | Out-Null
}

function RemoveSpaces {
  [CmdletBinding()]
  param (
    [Parameter(Position = 1)]
    [string] $Text
  )
  if ([string]::IsNullOrEmpty($Text)) {
    return [string]::Empty
  }

  return $Text.Trim()
}

function SplitAddress {
  [CmdletBinding()]
  param (
    [Parameter(Position = 1)]
    [string] $Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return [PSCustomObject]@{
      No  = ""
      Rue = ""
    }
  }

  "$($Text)" -match "(?<rue>(?:(?:^[A-Z])|(?:\s+[A-Z]))(?:.|\s)*)"
  
  $rue = RemoveSpaces $Matches.rue
  $no = RemoveSpaces ($Text.Substring(0, $Text.Length - $rue.Length))
  $result = [PSCustomObject]@{
    No  = $no
    Rue = $rue
  }
  return $result
}

function SplitMaterial {
  [CmdletBinding()]
  param (
    [Parameter(Position = 1)]
    $Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $Text
  }

  if ($Text -like "*Plomb*") {
    $result = "Plomb"
  }
  elseif ($Text -like "*Inconnu*" -or $Text -like "4*") {
    $result = "Inconnu"
  }
  elseif ($Text -like "*n/a*" -or $Text -clike "NA") {
    $result = "n/a"
  }
  else {
    $result = "Cuivre"
  }
  
  return $result
}

function GetTermGuidInCache {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, Position = 1)]
    $TermTable,
    [Parameter(Position = 2)]
    [string] $Term
  )

  if ([string]::IsNullOrWhiteSpace($Term)) {
    return [string]::Empty
  }
  
  $Term = RemoveSpaces $Term
  $result = $TermTable | Where-Object { $_.Name -eq "$($Term)" } | Select-Object Id

  if ([string]::IsNullOrWhiteSpace($result.Id)) {
    WriteErrorLog "Term" $Term
  }

  return "$($result.Id)"
}

function WriteErrorLog {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory, Position = 1)]
    [string] $Name,
    [Parameter(Position = 2)]
    $Value = "-",
    [switch] $Force)

  if (!($CheckOnly -or $Force)) {
    return 
  }
  $logFileName = Join-Path -Path "$($FilesFolder)" -ChildPath "_$($Name).log"
  "Secteur: '$($global:currentSecteur)' `tMslink: '$($global:currentMslink)' `t$($Name):'$($Value)'" >> "$($logFileName)"
}

function GetAllTerms {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, Position = 1)]
    [string] $TermSet,
    [string] $TermGroup
  )
  
  $result = @()

  if ([string]::IsNullOrWhiteSpace($TermGroup)) {
    $TermGroup = "TE"
  }

  $allTerms = Get-PnPTerm -TermGroup "$($TermGroup)" -TermSet "$($TermSet)"
 
  foreach ($term in $allTerms) {
    $withLabels = Get-PnPTerm -TermGroup "$($TermGroup)" -TermSet "$($TermSet)" -Identity "$($term.Name)" -Includes Labels
    foreach ($label in $withLabels.Labels) {
      $result += [PSCustomObject]@{
        Name = $label.Value
        Id   = $term.Id
      }
    }
  }
  
  return $result
}

function FillRequiredTerms {
  $global:zoneTerms += GetAllTerms "Zone Opération"
  $global:phaseTerms += GetAllTerms "Phase"
  $global:materialTerms += GetAllTerms "Matériau"
  $global:secteurs += GetAllSecteurs
}

function GetAllSecteurs {
  [CmdletBinding()]
  param (  )
  
  $secteurs = (Get-PnPListItem -List Secteurs -Fields Id, OrdrePriorisation).FieldValues
  $result = @()
  foreach ($s in $secteurs) {
    $result += [PSCustomObject]@{Ordre = $s.OrdrePriorisation; Id = $s.ID }
  }
  return $result
}

function GetNumberFrom {
  param (
    [Parameter(Position = 1)]
    [string] $Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return [string]::Empty
  }

  [ref]$result = 0.0

  $Text = RemoveSpaces $Text
  if ([float]::TryParse($Text, $result)) {
    return $result.Value
  }

  WriteErrorLog "Number" $Text
}

function GetDateFrom {
  param (
    [Parameter(Position = 1)]
    [string] $Text
  )

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return [string]::Empty
  }
 
  $Text = RemoveSpaces $Text
  $formats = "d-M-yyyy", "d/M/yyyy", "yyyy-M-d", "d-MMM.-yy", "yyyy-MMdd", "yyyyMMdd", "M-d-yyyy", "0dd-M-yyyy", "yyyy-M-0dd", "yyyy-d-M"
  [ref]$dateResult = get-date
  foreach ($f in $formats) {
    if ([datetime]::TryParseExact($Text, $f, $null, [System.Globalization.DateTimeStyles]::None, $dateResult)) {
      $result = "$($dateResult.Value.ToString('MM-dd-yyyy')) 13:00"
      return $result
    }
  }

  if ($Text -like "* et *") {
    $aux = $Text.Split("et")
    return GetDateFrom $aux[1]
  }

  WriteErrorLog "Date" $Text
  return [string]::Empty
}

function GetBooleanFrom {
  param (
    [Parameter(Position = 1)]
    [string] $Text
  )

  $Text = RemoveSpaces $Text
  $result = !([string]::IsNullOrWhiteSpace($Text) -or $Text -like "*non*" -or $Text -like "0" -or $Text -like "*n/a*")
  if ($result) {
    WriteErrorLog "YesNo" $Text
  }
  return $result
}

function ExtractExcelInformations {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, Position = 1)]
    [string] $FileName
  )

  $objExcel = New-Excel -Path "$($FileName)"
  $Worksheet = $objExcel.Workbook.Worksheets[1]
  if(!$?){
    $Worksheet = $objExcel | Get-Worksheet -Name "Adresses"
  }
  $totalNoOfRecords = $Worksheet.Dimension.Rows
  $totalNoOfItems = $totalNoOfRecords - 1
  $startRow = 3
  $msLinkColumn = 2
  $addressColumn = 3
  $anneeColumn = 4
  $commentairesColumn = 5
  $zoneColumn = 6
  $phaseColumn = 7
  $dateCaractVilleColumn = 8
  $materialVdqColumn = 9
  $longueurVilleColumn = 10
  $dateRemplacementVilleColumn = 11
  $dateCaractPriveColumn = 12
  $materialPriveeColumn = 13
  $longueurPriveColumn = 14
  $dateRemplacementPriveColumn = 15
  $proprietaireColumn = 16
  $noTelColumn = 17
  $emailColumn = 18
  $adresseProprietaireColumn = 19
  $remisePichetColumn = 23
  $arbreMatureAConsidererColumn = 27
  $potentielArchealogiqueColumn = 28
  $subventionColumn = 30

  for ($i = 0; $i -le $totalNoOfItems; $i++) {
    $global:currentMslink = RemoveSpaces $($Worksheet.Cells.Item($startRow + $i, $msLinkColumn).text)
    if ([string]::IsNullOrEmpty($global:currentMslink)) {
      continue
    }
    $secteur = GetCurrentSecteurInCache
    $fullAddress = SplitAddress $($Worksheet.Cells.Item($startRow + $i, $addressColumn).text)
    $noCivique = $fullAddress.No
    $address = $fullAddress.Rue
    $annee = GetNumberFrom $($Worksheet.Cells.Item($startRow + $i, $anneeColumn).text)
    $commentaires = RemoveSpaces $($Worksheet.Cells.Item($startRow + $i, $commentairesColumn).text)
    $zone = GetTermGuidInCache $global:zoneTerms "$($Worksheet.Cells.Item($startRow + $i, $zoneColumn).text)"
    $phase = GetTermGuidInCache $global:phaseTerms "$($Worksheet.Cells.Item($startRow + $i, $phaseColumn).text)"
    $dateCaractVille = GetDateFrom "$($Worksheet.Cells.Item($startRow + $i, $dateCaractVilleColumn).text)"
    $matVdq = SplitMaterial "$($Worksheet.Cells.Item($startRow + $i, $materialVdqColumn).text)"
    $materialVdq = GetTermGuidInCache $global:materialTerms "$($matVdq)"
    $longueurVille = GetNumberFrom $($Worksheet.Cells.Item($startRow + $i, $longueurVilleColumn).text)
    $dateRemplacementVille = GetDateFrom "$($Worksheet.Cells.Item($startRow + $i, $dateRemplacementVilleColumn).text)"
    $dateCaractPrive = GetDateFrom "$($Worksheet.Cells.Item($startRow + $i, $dateCaractPriveColumn).text)"
    $matPriv = SplitMaterial "$($Worksheet.Cells.Item($startRow + $i, $materialPriveeColumn).text)"
    $materialPrivee = GetTermGuidInCache $global:materialTerms "$($matPriv)"
    $longueurPrive = GetNumberFrom "$($Worksheet.Cells.Item($startRow + $i, $longueurPriveColumn).text)"
    $dateRemplacementPrive = GetDateFrom "$($Worksheet.Cells.Item($startRow + $i, $dateRemplacementPriveColumn).text)"
    $proprietaire = RemoveSpaces $($Worksheet.Cells.Item($startRow + $i, $proprietaireColumn).text)
    $noTel = RemoveSpaces $($Worksheet.Cells.Item($startRow + $i, $noTelColumn).text)
    $email = RemoveSpaces $($Worksheet.Cells.Item($startRow + $i, $emailColumn).text)
    $adresseProprietaire = RemoveSpaces $($Worksheet.Cells.Item($startRow + $i, $adresseProprietaireColumn).text)
    $remisePichet = GetBooleanFrom $($Worksheet.Cells.Item($startRow + $i, $remisePichetColumn).text)
    $arbreMatureAConsiderer = GetBooleanFrom $($Worksheet.Cells.Item($startRow + $i, $arbreMatureAConsidererColumn).text)
    $potentielArchealogique = GetBooleanFrom $($Worksheet.Cells.Item($startRow + $i, $potentielArchealogiqueColumn).text)
    $subvention = GetBooleanFrom $($Worksheet.Cells.Item($startRow + $i, $subventionColumn).text)

    $it = @{}
    $it = AddProperty $it "Secteur" $secteur
    $it = AddProperty $it "Mslink" $global:currentMslink
    $it = AddProperty $it "NoCivique" $noCivique
    $it = AddProperty $it "Rue" $address
    $it = AddProperty $it "AnneeConstruction" $annee
    $it = AddProperty $it "Commentaires" $commentaires
    $it = AddProperty $it "ZoneOperationTPPropriete" $zone
    $it = AddProperty $it "Phase" $phase
    $it = AddProperty $it "DateCaractVille" $dateCaractVille
    $it = AddProperty $it "MateriauVille" $materialVdq
    $it = AddProperty $it "LongueurVille" $longueurVille
    $it = AddProperty $it "DateRemplacementVille" $dateRemplacementVille
    $it = AddProperty $it "DateCaractPrive" $dateCaractPrive
    $it = AddProperty $it "MateriauPrive" $materialPrivee
    $it = AddProperty $it "LongueurPrive" $longueurPrive
    $it = AddProperty $it "DateRemplacementPrive" $dateRemplacementPrive
    $it = AddProperty $it "Proprietaire" $proprietaire
    $it = AddProperty $it "NoTelephoneProprietaire" $noTel
    $it = AddProperty $it "CourrielProprietaire" $email
    $it = AddProperty $it "AdresseProprietaire" $adresseProprietaire
    $it = AddProperty $it "RemisePichet" $remisePichet
    $it = AddProperty $it "ArbreMatureAConsiderer" $arbreMatureAConsiderer
    $it = AddProperty $it "PotentielArchealogique" $potentielArchealogique
    $it = AddProperty $it "Subvention" $subvention

    $global:allValues += [PSCustomObject]@{
      Item = $it
    }
  }
}

function GetCurrentSecteurInCache {
  [CmdletBinding()]
  param ()

  $result = $global:secteurs | Where-Object { $_.Ordre -eq $global:currentSecteur } 
  return $result.Id
}

function ReadAllFiles {
  [CmdletBinding()]
  param(
    [Parameter(Position = 1)]
    [string] $FilesFolder,
    [Parameter(Position = 2)]
    [string] $FilesFilter
  )

  $files = Get-ChildItem -Path "$($FilesFolder)" -Filter "$($FilesFilter)" -File -Name
  foreach ($file in $files) {
    $aux = $file.Split("_")
    $global:currentSecteur = [int]$aux[1]
    if ($global:currentSecteur -gt 0) {
      $path = Join-Path -Path "$($FilesFolder)" -ChildPath "$($file)"
      ExtractExcelInformations "$($path)"
    }
  }
}

function CreateAllDocumentset {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $True, Position = 1)]
    [string] $ListName,
    [Parameter(Mandatory = $True, Position = 2)]
    [string] $DocumentSetName
  )

  if ($CheckOnly) {
    return
  }

  foreach ($v in $global:allValues) {
    $val = $v.Item
    $global:currentSecteur = $val.Secteur
    $global:currentMslink = $val.Mslink
    if (!$UpdateOnly) {
      $success = $False
      for ($i = 0; !$success -and $i -lt 3; $i++) {
        if ($i -gt 0) {
          Write-Host "Retrying to create document: $($global:currentMslink) Secteur: '$($global:currentSecteur)'" -ForegroundColor Cyan
        }
        Add-PnPDocumentSet -List "$($ListName)" -ContentType "$($DocumentSetName)" -Name "$($global:currentMslink)" | Out-Null
        $success = $?
      }
      if (!$success) {
        WriteErrorLog "NoCreated" -Force
        continue
      }
    }

    if (!$CreateOnly) {
      $success = $False
      for ($i = 0; !$success -and $i -lt 3; $i++) {
        if ($i -gt 0) {
          Write-Host "Retrying to get current document $($global:currentMslink) Secteur: '$($global:currentSecteur)'" -ForegroundColor Cyan
        }
        $query = "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$global:currentMslink</Value></Eq></Where></Query></View>"
        $item = Get-PnPListItem -List "$($ListName)" -Query "$($query)"
        $success = $?
      }
      if (!$success) {
        WriteErrorLog "NoRetrivied" -Force
        continue
      }
      $success = $False
      for ($i = 0; !$success -and $i -lt 3; $i++) {
        if ($i -gt 0) {
          Write-Host "Retrying to update document $($global:currentMslink) Secteur: '$($global:currentSecteur)'" -ForegroundColor Cyan
        }
        Set-PnPListItem -List "$($ListName)" -Identity $item -Values $val | Out-Null
        $success = $?
      }
      if (!$success) {
        WriteErrorLog "NoUpdated" -Force
      }
    }
  }
}

function AddProperty {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory, Position = 1)]
    $HastTable,
    [Parameter(Mandatory, Position = 2)]
    [string] $PropertyName,
    [Parameter(Position = 3)]
    $PropertyValue
  )

  if (![string]::IsNullOrWhiteSpace($PropertyValue)) {
    $HastTable += @{ $PropertyName = $PropertyValue }
  }
  return $HastTable
}

$global:allValues = @()
$global:zoneTerms = @()
$global:phaseTerms = @()
$global:materialTerms = @()
$global:secteurs = @()
$global:currentSecteur = 0
$global:currentMslink = 0

InstallPSExcel
ConnectToHost "$($SiteUrl)"
FillRequiredTerms
ReadAllFiles "$($FilesFolder)" "$($FilesFilter)"
CreateAllDocumentset "$($ListName)" "$($DocumentSetName)"
DisconnectHost