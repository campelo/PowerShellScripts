function String-ToAlphaNumeric {
  <#
.SYNOPSIS
  This function will remove the diacritics (accents) characters from a string.

.DESCRIPTION
  This function will remove the diacritics (accents) characters from a string.

.PARAMETER String
  Specifies the String(s) on which the diacritics need to be removed

.PARAMETER NormalizationForm
  Specifies the normalization form to use
  https://msdn.microsoft.com/en-us/library/system.text.normalizationform(v=vs.110).aspx

.EXAMPLE
  PS C:\> Remove-StringDiacritic "L'été de Raphaël"
  LetedeRaphael
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

  Write-Verbose -Message "$MainString"
  try {
    # Normalize the String
    $Normalized = $MainString.Normalize($NormalizationForm)
    $NewString = New-Object -TypeName System.Text.StringBuilder

    # Convert the String to CharArray
    $normalized.ToCharArray() |
    ForEach-Object -Process {
      if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -eq [Globalization.UnicodeCategory]::LowercaseLetter -or [Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -eq [Globalization.UnicodeCategory]::UppercaseLetter -or [Globalization.CharUnicodeInfo]::GetUnicodeCategory($psitem) -eq [Globalization.UnicodeCategory]::DecimalDigitNumber) {
        [void]$NewString.Append($psitem)
      }
    }

    #Combine the new string chars
    write-host "$($NewString -as [string])" -foregroundcolor Green 
  }
  Catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
  }
}