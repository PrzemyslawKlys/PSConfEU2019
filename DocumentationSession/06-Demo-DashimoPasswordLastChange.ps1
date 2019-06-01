Import-Module PSWinDocumentation.AD -Force
Import-Module Dashimo

if ($null -eq $DataSetForest) {
    $DataSetForest = Get-WinADForestInformation -Verbose -PasswordQuality -DontRemoveEmpty -Splitter "`r`n"
}

<# Single, named domain
Dashboard -Name 'Dashimo Test' -FilePath "$PSScriptRoot\06-Demo-01.html" -Show {
    Section -Name 'Domain Users' {
        Table -DataTable $DataSetForest.FoundDomains.'ad.evotec.xyz'.DomainUsers {
            TableConditionalFormatting -Name 'PasswordLastChanged(Days)' -ComparisonType number -Operator gt -Value 90 -Row -BackgroundColor RosyBrown
        }
    }
}
#>

## Multi Forest / Dynamic domain

Dashboard -Name 'Dashimo Test' -FilePath "$PSScriptRoot\06-Demo-02.html" -Show {
    foreach ($Domain in $DataSetForest.FoundDomains.Keys) {
        Section -Name "Domain Users $Domain" {
            Table -DataTable $DataSetForest.FoundDomains.$Domain.DomainUsers {
                TableConditionalFormatting -Name 'PasswordLastChanged(Days)' -ComparisonType number -Operator gt -Value 30 -BackgroundColor Beige
                TableConditionalFormatting -Name 'PasswordLastChanged(Days)' -ComparisonType number -Operator gt -Value 90 -BackgroundColor RosyBrown
                TableConditionalFormatting -Name 'PasswordLastChanged(Days)' -ComparisonType number -Operator lt -Value 30 -BackgroundColor Aquamarine
                TableConditionalFormatting -Name 'PasswordLastChanged(Days)' -ComparisonType number -Operator lt -Value 10 -BackgroundColor Green
            }
        }
    }
}