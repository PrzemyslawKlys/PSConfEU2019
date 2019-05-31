Import-Module PSWinDocumentation.AD.psd1 -Force

#if ($null -eq $DataSetForest) {
    $DataSetForest = Get-WinADForestInformation -Verbose -PasswordQuality -DontRemoveEmpty -Splitter "`r`n"
#}
$DataSetForest