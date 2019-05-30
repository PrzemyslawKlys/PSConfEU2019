
$InvoiceEntry7 = [ordered]@{}
$InvoiceEntry7.Description = 'IT Services 4'
$InvoiceEntry7.Amount = '$301'

$InvoiceEntry8 = [ordered]@{}
$InvoiceEntry8.Description = 'IT Services 5'
$InvoiceEntry8.Amount = '$299'

$InvoiceDataOrdered1 = @(
    $InvoiceEntry7
)
$InvoiceDataOrdered2 = @(
    $InvoiceEntry7
    $InvoiceEntry8
)


Import-Module PSWriteWord #-Force
$FilePath = "$Env:USERPROFILE\Desktop\PSWriteWord-Example-Tables11.docx"

$WordDocument = New-WordDocument $FilePath
Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceDataOrdered1 -Design ColorfulGrid -Supress $true -Transpose
Add-WordParagraph -WordDocument $WordDocument -Supress $True

Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceDataOrdered2 -Design ColorfulGrid -Percentage $true -Transpose -Supress $True

Save-WordDocument $WordDocument -Language 'en-US' -Supress $True -OpenDocument