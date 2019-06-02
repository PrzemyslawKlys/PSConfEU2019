Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

$FilePathWord = "$PSScriptRoot\01-Demo.docx"
$FilePathPDF = "$PSScriptRoot\01-Demo.pdf"

$word = New-Object -ComObject word.application
$word.Visible = $true

$doc = $word.documents.add()
$doc.Styles["Normal"].ParagraphFormat.SpaceAfter = 0
$doc.Styles["Normal"].ParagraphFormat.SpaceBefore = 0
$margin = 36 # 1.26 cm
$doc.PageSetup.LeftMargin = $margin
$doc.PageSetup.RightMargin = $margin
$doc.PageSetup.TopMargin = $margin
$doc.PageSetup.BottomMargin = $margin
$selection = $word.Selection

$selection.Style = "Heading 1"
$selection.TypeText('Test1')
$selection.TypeParagraph()
$selection.Style = "Normal"
$selection.TypeText('Test2')
$selection.TypeParagraph()
$selection.TypeText('test3')

$doc.SaveAs([ref]$FilePathWord, [ref] [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault)
$doc.SaveAs([ref]$FilePathPDF, [ref] [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF)

#$doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
#$word.Quit()