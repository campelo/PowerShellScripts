##function String-ToAlphaNumeric {
  <#
.SYNOPSIS
  This function will remove the diacritics (accents) characters from a string.

.DESCRIPTION
  This function will remove the diacritics (accents) characters from a string.

.PARAMETER MainString
  Specifies the String(s) on which the diacritics need to be removed

.PARAMETER NormalizationForm
  Specifies the normalization form to use
  https://msdn.microsoft.com/en-us/library/system.text.normalizationform(v=vs.110).aspx

.EXAMPLE
  PS C:\> String-ToAlphaNumeric "L'été de Raphaël-autre"
  LEteDeRaphaelAutre
.NOTES
  
#>
  [CmdletBinding()]
  PARAM
  (
    [ValidateNotNullOrEmpty()]
    [Alias('Text')]
    [System.String[]]$MainString,
    [System.Text.NormalizationForm]$NormalizationForm = "FormD"
  )

  $result
  $toUpper = $true
  try {
    # Normalize the String
    $Normalized = $MainString.Normalize($NormalizationForm)
    $NewString = New-Object -TypeName System.Text.StringBuilder

    # Convert the String to CharArray
    $normalized.ToCharArray() |
    ForEach-Object -Process {
      if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -eq [Globalization.UnicodeCategory]::LowercaseLetter -or [Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -eq [Globalization.UnicodeCategory]::UppercaseLetter -or [Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -eq [Globalization.UnicodeCategory]::DecimalDigitNumber) {
		if($toUpper){
			[void]$NewString.Append(($psitem -as [string]).ToUpper())
			$toUpper = $false
		} else {
			[void]$NewString.Append($psitem)
		}
      } elseif ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
		$toUpper = $true
	  }
    }

    #Combine the new string chars
	$result = $($NewString -as [string])
    write-Verbose "$($result)" # -foregroundcolor Green 
  }
  Catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
  }
  return $($result)
##}