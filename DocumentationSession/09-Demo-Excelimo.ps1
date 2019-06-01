﻿Import-Module PSWinDocumentation.AD -Force
Import-Module Excelimo -Force

if ($null -eq $ADForest) {
    $ADForest = Get-WinADForestInformation -Verbose -PasswordQuality
}

Excel -FilePath $PSScriptRoot\"09-Demo.xlsx" {
    WorkbookProperties -Title 'PSWinDocumentation - Active Directory Demo' -Author 'Przemysław Kłys'

    Worksheet -DataTable $ADForest.ForestInformation -Name 'Forest Information' -TabColor Green -AutoFilter -AutoFit
    Worksheet -DataTable $ADForest.ForestSites -Name 'Forest Sites' -TabColor RoyalBlue -AutoFilter -AutoFit
    Worksheet -DataTable $ADForest.ForestOptionalFeatures -Name 'Forest Optional Features' -TabColor Red -AutoFilter -AutoFit
} -Open