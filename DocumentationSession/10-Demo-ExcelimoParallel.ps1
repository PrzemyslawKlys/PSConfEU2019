Import-Module Excelimo -Force

$Process0 = Get-Process | Select-Object -First 1
$Process1 = Get-Process | Select-Object -First 5
$Process2 = Get-Process | Select-Object -First 10
$Process3 = Get-Process | Select-Object -First 15
$Process4 = Get-Process | Select-Object -First 20
$Process5 = Get-Process | Select-Object -First 25
$Process6 = Get-Process | Select-Object -First 30
$Process7 = Get-Process | Select-Object -First 35
$Process8 = Get-Process | Select-Object -First 40
$Process9 = Get-Process | Select-Object -First 45
$Process10 = Get-Process | Select-Object -First 50

Excel -FilePath $PSScriptRoot\"10-Demo.xlsx" {
    WorkbookProperties -Title 'Test'
    Worksheet -DataTable $Process0 -Name 'Processes'

    Worksheet -DataTable $Process1 -Name 'Processes Test1' -TabColor Crimson
    Worksheet -DataTable $Process2 -Name 'Processes Test2' -TabColor BlueViolet
    Worksheet -DataTable $Process3 -Name 'Processes Test3' -TabColor Aquamarine
    Worksheet -DataTable $Process4 -Name 'Processes Test4' -TabColor LemonChiffon
    Worksheet -DataTable $Process5 -Name 'Processes Test5' -TabColor Blue
    Worksheet -DataTable $Process6 -Name 'Processes Test6' -TabColor Bisque
    Worksheet -DataTable $Process7 -Name 'Processes Test7' -TabColor DodgerBlue
    Worksheet -DataTable $Process8 -Name 'Processes Test8' -TabColor Aqua
    Worksheet -DataTable $Process9 -Name 'Processes Test9' -TabColor DarkCyan
    Worksheet -DataTable $Process10 -Name 'Processes Test10' -TabColor Red
} -Verbose -Open #-Parallel