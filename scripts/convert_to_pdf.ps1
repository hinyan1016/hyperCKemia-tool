$pptxPath = (Resolve-Path 'C:\Users\jsber\OneDrive\Documents\Claude_task\indian_restaurant_slides.pptx').Path
$pdfPath = $pptxPath -replace '\.pptx$','.pdf'
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $ppt.Presentations.Open($pptxPath)
$presentation.SaveAs($pdfPath, 32)
$presentation.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
Write-Host 'PDF created successfully'
