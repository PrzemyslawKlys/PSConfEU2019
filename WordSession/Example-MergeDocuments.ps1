Import-Module PSWriteWord -Force

$FilePath1 = "$PSScriptRoot\02-Demo.docx"
$FilePath2 = "$PSScriptRoot\03-Demo.docx"
$FileOutput = "$PSScriptRoot\11-Demo-MergedDocument.docx"

Merge-WordDocument -FilePath1 $FilePath1 -FilePath2 $FilePath2 -FileOutput $FileOutput -Supress $true