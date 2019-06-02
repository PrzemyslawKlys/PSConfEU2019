code C:\Users\przemyslaw.klys.SUPPORT\Documents\PowerShell\Modules\PSWriteWord\1.0.0\PSWriteWord.psm1

<#
if ($PSEdition -eq 'Core') {
    Add-Type -Path $PSScriptRoot\Lib\Core\Xceed.Document.NETStandard.dll
    Add-Type -Path $PSScriptRoot\Lib\Core\Xceed.PDF.NETStandard.dll
    Add-Type -Path $PSScriptRoot\Lib\Core\Xceed.Words.NETStandard.dll
} else { Add-Type -Path $PSScriptRoot\Lib\Default\Xceed.Words.NET.dll }
. $PSScriptRoot\PSWriteWord.ps1
#>