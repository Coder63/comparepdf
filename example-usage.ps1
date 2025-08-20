# Example usage script for Compare-PDFs.ps1
# This demonstrates different ways to use the PDF comparison script

Write-Host "=== PDF Comparison Script Examples ===" -ForegroundColor Cyan

# Example 1: Basic comparison with sample files
Write-Host "`nExample 1: Basic PDF Comparison" -ForegroundColor Yellow
Write-Host ".\Compare-PDFs.ps1 -FirstPDF 'C:\Documents\original.pdf' -SecondPDF 'C:\Documents\revised.pdf' -OutputDirectory 'C:\Reports'"

# Example 2: Custom report name
Write-Host "`nExample 2: Custom Report Name" -ForegroundColor Yellow
Write-Host ".\Compare-PDFs.ps1 -FirstPDF '.\version1.pdf' -SecondPDF '.\version2.pdf' -OutputDirectory '.\results' -ReportName 'Contract_Review_v1_vs_v2'"

# Example 3: Batch processing function
function Compare-PDFBatch {
    param(
        [array]$PDFPairs,
        [string]$BaseOutputDir
    )
    
    foreach ($pair in $PDFPairs) {
        $outputDir = Join-Path $BaseOutputDir $pair.Name
        Write-Host "Processing: $($pair.Name)" -ForegroundColor Green
        
        try {
            .\Compare-PDFs.ps1 -FirstPDF $pair.First -SecondPDF $pair.Second -OutputDirectory $outputDir -ReportName $pair.Name
        }
        catch {
            Write-Host "Failed to process $($pair.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Example batch data
$sampleBatch = @(
    @{Name="Contract_Changes"; First="contract_v1.pdf"; Second="contract_v2.pdf"},
    @{Name="Proposal_Updates"; First="proposal_draft.pdf"; Second="proposal_final.pdf"},
    @{Name="Manual_Revisions"; First="manual_old.pdf"; Second="manual_new.pdf"}
)

Write-Host "`nExample 3: Batch Processing" -ForegroundColor Yellow
Write-Host "Compare-PDFBatch -PDFPairs `$sampleBatch -BaseOutputDir 'C:\BatchReports'"

# Example 4: Error handling wrapper
function Compare-PDFsWithRetry {
    param(
        [string]$FirstPDF,
        [string]$SecondPDF,
        [string]$OutputDirectory,
        [int]$MaxRetries = 3
    )
    
    $attempt = 1
    while ($attempt -le $MaxRetries) {
        try {
            Write-Host "Attempt $attempt of $MaxRetries" -ForegroundColor Yellow
            $result = .\Compare-PDFs.ps1 -FirstPDF $FirstPDF -SecondPDF $SecondPDF -OutputDirectory $OutputDirectory
            return $result
        }
        catch {
            Write-Host "Attempt $attempt failed: $($_.Exception.Message)" -ForegroundColor Red
            $attempt++
            if ($attempt -le $MaxRetries) {
                Start-Sleep -Seconds 5
            }
        }
    }
    throw "All $MaxRetries attempts failed"
}

Write-Host "`nExample 4: Retry Logic" -ForegroundColor Yellow
Write-Host "Compare-PDFsWithRetry -FirstPDF 'doc1.pdf' -SecondPDF 'doc2.pdf' -OutputDirectory 'output' -MaxRetries 3"

Write-Host "`n=== Ready to use! ===" -ForegroundColor Green
Write-Host "Copy and modify these examples for your specific needs." -ForegroundColor White
