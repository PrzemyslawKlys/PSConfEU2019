﻿Import-Module PSWriteWord -Force

$FilePath = "$PSScriptRoot\02-Demo-01.docx"

### define new document
$WordDocument = New-WordDocument $FilePath -Verbose
### add 3 paragraphs
Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
$Paragraph = Add-WordPageBreak -WordDocument $WordDocument -Verbose
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21 -Supress $True
$Paragraph = Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $false
$Paragraph | Add-WordPageBreak -InsertWhere BeforeSelf -Supress $True -Verbose

### Save document
Save-WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument