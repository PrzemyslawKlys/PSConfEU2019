Import-Module PSWriteWord -Force

$FilePath = "$PSScriptRoot\07-Demo-01.docx"

$myitems = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)

$WordDocument = New-WordDocument $FilePath

Add-WordTable -WordDocument $WordDocument -DataTable $myitems -Design ColorfulList -Supress $True #-Verbose

Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $myitems1 -Design ColorfulList -Supress $True #-Verbose

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument