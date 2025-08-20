# Test script for Compare-PDFs-Acrobat.ps1
# This demonstrates how to use the Acrobat native comparison script

Write-Host "=== Testing Acrobat PDF Comparison Script ===" -ForegroundColor Cyan

# Example 1: Basic usage
Write-Host "`nExample 1: Basic PDF Comparison" -ForegroundColor Yellow
Write-Host "Command: .\Compare-PDFs-Acrobat.ps1 -FirstPDF './one/file1.pdf' -SecondPDF './two/file1.pdf' -OutputDirectory './pdf-reports'"

# Example 2: Custom report name
Write-Host "`nExample 2: Custom Report Name" -ForegroundColor Yellow
Write-Host "Command: .\Compare-PDFs-Acrobat.ps1 -FirstPDF 'doc1.pdf' -SecondPDF 'doc2.pdf' -OutputDirectory 'C:\Reports' -ReportName 'Contract_Changes'"

# Example 3: Batch processing
Write-Host "`nExample 3: Batch Processing" -ForegroundColor Yellow
Write-Host @"
`$pdfPairs = @(
    @{First='./one/file1.pdf'; Second='./two/file1.pdf'; Name='FileComparison1'},
    @{First='doc1.pdf'; Second='doc2.pdf'; Name='FileComparison2'}
)

foreach (`$pair in `$pdfPairs) {
    .\Compare-PDFs-Acrobat.ps1 -FirstPDF `$pair.First -SecondPDF `$pair.Second -OutputDirectory './batch-reports' -ReportName `$pair.Name
}
"@

Write-Host "`n=== Key Features ===" -ForegroundColor Green
Write-Host "✅ Uses Acrobat's native comparePages() JavaScript API"
Write-Host "✅ Generates PDF comparison reports (not text files)"
Write-Host "✅ Visual highlighting of differences"
Write-Host "✅ Automatic fallback mechanisms"
Write-Host "✅ Works on Windows with Adobe Acrobat Pro"

Write-Host "`n=== Requirements ===" -ForegroundColor Red
Write-Host "⚠️  Adobe Acrobat Pro DC (not Reader)"
Write-Host "⚠️  Windows OS (COM interface required)"
Write-Host "⚠️  PowerShell 5.1 or later"

Write-Host "`n=== Current Status ===" -ForegroundColor White
Write-Host "Script Location: Compare-PDFs-Acrobat.ps1"
Write-Host "Documentation: README-Acrobat-Native.md"
Write-Host "Test Files Available: ./one/file1.pdf, ./two/file1.pdf"

Write-Host "`nReady to use on Windows systems with Adobe Acrobat Pro!" -ForegroundColor Green
