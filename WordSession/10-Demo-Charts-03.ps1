﻿Import-Module PSWriteWord #-Force

$FilePath = "$PSScriptRoot\10-Demo-03.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #1' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances' -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000 -ChartLegendPosition Bottom -ChartLegendOverlay $false

Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #2' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True

$Series1 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series2 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series1, $Series2

Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #3' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True


$Series3 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series4 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000


Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series3, $Series4 -BarGrouping Stacked

Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #4' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True


$Series5 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series6 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series5, $Series6 -BarDirection Column


Add-WordText -WordDocument $WordDocument -Text 'Bar Chart Example #3 - No Legend' `
    -FontSize 15 `
    -Color Blue `
    -Bold $true -HeadingType Heading1 -Supress $True

$Series7 = Add-WordChartSeries -ChartName 'One'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 1050.50, 2000, 20000
$Series8 = Add-WordChartSeries -ChartName 'Two'  -Names 'Today', 'Yesterday', 'Two days ago' -Values 3000, 2000, 1000

Add-WordBarChart -WordDocument $WordDocument -ChartName 'My finances'-ChartLegendPosition Bottom -ChartLegendOverlay $false -ChartSeries $Series7, $Series8 -BarDirection Column -NoLegend

Save-WordDocument $WordDocument -Supress $True -OpenDocument