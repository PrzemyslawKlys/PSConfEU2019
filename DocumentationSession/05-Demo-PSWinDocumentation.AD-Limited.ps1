Import-Module PSWinDocumentation.AD.psd1 -Force

#[PSWinDocumentation.ActiveDirectory].GetEnumValues()

$TypesRequired = @(
    'DomainInformation'
    'ForestDomainControllers'
    'DomainUsers'
)

$DataSetLimited = Get-WinADForestInformation -Verbose -PasswordQuality -DontRemoveEmpty -Splitter "`r`n" -TypesRequired $TypesRequired
$DataSetLimited