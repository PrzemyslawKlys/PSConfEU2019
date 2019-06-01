Import-Module PSWinDocumentation.AD -Force
Import-Module Dashimo

if ($null -eq $DataSetForest) {
    $DataSetForest = Get-WinADForestInformation -Verbose -PasswordQuality -DontRemoveEmpty -Splitter "`r`n"
}

<# Not working due to Hashtable
Dashboard -Name 'Dashimo Test' -FilePath "$PSScriptRoot\06-Demo-02.html" -Show {
    foreach ($Domain in $DataSetForest.FoundDomains.Keys) {
        Section -Name "Domain Users $Domain" {
            Table -DataTable $DataSetForest.FoundDomains.$Domain.DomainDefaultPasswordPolicy {
                TableConditionalFormatting -Name 'Max Password Age' -ComparisonType number -Operator gt -Value 50 -Row -BackgroundColor RosyBrown
                TableConditionalFormatting -Name 'Max Password Age' -ComparisonType number -Operator lt -Value 50 -Row -BackgroundColor Green
            }
        }
    }
}
#>

Dashboard -Name 'Dashimo Test' -FilePath "$PSScriptRoot\06-Demo-02.html" -Show {
    foreach ($Domain in $DataSetForest.FoundDomains.Keys) {
        Section -Name "Domain Users $Domain" {

            $PassswordPolicyTransposed = Format-TransposeTable -Object $DataSetForest.FoundDomains.$Domain.DomainDefaultPasswordPolicy

            Table -DataTable $PassswordPolicyTransposed {
                TableConditionalFormatting -Name 'Max Password Age' -ComparisonType number -Operator gt -Value 50 -BackgroundColor RosyBrown
                TableConditionalFormatting -Name 'Max Password Age' -ComparisonType number -Operator lt -Value 50 -BackgroundColor Green
            }
        }
    }
}