Add-Type -AssemblyName "Microsoft.Office.Interop.Word"

$doc.SaveAs([ref]$outputPath, [ref] [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF)

$doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)

$word.Quit()