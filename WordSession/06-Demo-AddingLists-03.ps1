
Import-Module PSWriteWord -Force

$FilePath = "$PSScriptRoot\06-Demo-03.docx"

### define new document
$WordDocument = New-WordDocument $FilePath -Verbose

Add-WordText -WordDocument $WordDocument -Text 'This is a text' -FontSize 10 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 21' -FontSize 21 -Supress $True
Add-WordText -WordDocument $WordDocument -Text 'This is a text font size 15' -FontSize 15 -Supress $True

New-WordList -WordDocument $WordDocument {
    New-WordListItem -Level 0 -Text 'Test'
    New-WordListItem -Level 1 -Text 'Test1'
    New-WordListItem -Level 0 -Text 'Test2'
}

Save-WordDocument $WordDocument -Language 'en-US' -Supress $true -OpenDocument