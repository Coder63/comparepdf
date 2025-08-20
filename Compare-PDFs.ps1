param(
    [Parameter(Mandatory=$true)]
    [string]$FirstPDF,
    
    [Parameter(Mandatory=$true)]
    [string]$SecondPDF,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputDirectory,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportName = "PDF_Comparison_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
)

# Function to validate Adobe Acrobat installation
function Test-AcrobatInstallation {
    try {
        $acroApp = New-Object -ComObject AcroExch.App
        $acroApp.Exit()
        return $true
    }
    catch {
        return $false
    }
}

# Function to validate PDF files
function Test-PDFFile {
    param([string]$FilePath)
    
    if (-not (Test-Path $FilePath)) {
        throw "PDF file not found: $FilePath"
    }
    
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    if ($extension -ne ".pdf") {
        throw "File is not a PDF: $FilePath"
    }
    
    return $true
}

# Function to create output directory
function New-OutputDirectory {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
            Write-Host "Created output directory: $Path" -ForegroundColor Green
        }
        catch {
            throw "Failed to create output directory: $Path. Error: $($_.Exception.Message)"
        }
    }
}

# Function to compare PDFs using Acrobat
function Compare-PDFsWithAcrobat {
    param(
        [string]$PDF1,
        [string]$PDF2,
        [string]$OutputDir,
        [string]$ReportFileName
    )
    
    $acroApp = $null
    $avDoc1 = $null
    $avDoc2 = $null
    
    try {
        Write-Host "Initializing Adobe Acrobat..." -ForegroundColor Yellow
        
        # Create Acrobat application object
        $acroApp = New-Object -ComObject AcroExch.App
        $acroApp.Show()
        
        # Open first PDF
        Write-Host "Opening first PDF: $PDF1" -ForegroundColor Yellow
        $avDoc1 = New-Object -ComObject AcroExch.AVDoc
        if (-not $avDoc1.Open($PDF1, "")) {
            throw "Failed to open first PDF: $PDF1"
        }
        
        # Open second PDF
        Write-Host "Opening second PDF: $PDF2" -ForegroundColor Yellow
        $avDoc2 = New-Object -ComObject AcroExch.AVDoc
        if (-not $avDoc2.Open($PDF2, "")) {
            throw "Failed to open second PDF: $PDF2"
        }
        
        # Get PDDoc objects
        $pdDoc1 = $avDoc1.GetPDDoc()
        $pdDoc2 = $avDoc2.GetPDDoc()
        
        # Prepare comparison report path
        $reportPath = Join-Path $OutputDir "$ReportFileName.pdf"
        
        Write-Host "Performing PDF comparison..." -ForegroundColor Yellow
        
        # Use Acrobat's compare documents feature
        # This requires Acrobat Pro with comparison capabilities
        $jsObject = $pdDoc1.GetJSObject()
        
        # JavaScript code for comparison
        $jsCode = @"
var oDoc = this;
var compareDoc = app.openDoc('$($PDF2.Replace('\', '\\'))');
var result = oDoc.comparePages({
    cOtherDoc: compareDoc,
    nStart: 0,
    nEnd: oDoc.numPages - 1,
    cUIPolicy: 'never'
});

if (result) {
    var reportDoc = result.document;
    reportDoc.saveAs('$($reportPath.Replace('\', '\\'))');
    reportDoc.closeDoc();
}
compareDoc.closeDoc();
"@
        
        try {
            $jsObject.eval($jsCode)
            Write-Host "Comparison completed successfully!" -ForegroundColor Green
            Write-Host "Report saved to: $reportPath" -ForegroundColor Green
        }
        catch {
            # Fallback: Create a basic comparison report
            Write-Host "Advanced comparison failed, creating basic report..." -ForegroundColor Yellow
            Create-BasicComparisonReport -PDF1 $PDF1 -PDF2 $PDF2 -OutputPath $reportPath
        }
        
        return $reportPath
    }
    catch {
        throw "PDF comparison failed: $($_.Exception.Message)"
    }
    finally {
        # Clean up COM objects
        if ($avDoc2) {
            try { $avDoc2.Close($true) } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc2) | Out-Null
        }
        if ($avDoc1) {
            try { $avDoc1.Close($true) } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc1) | Out-Null
        }
        if ($acroApp) {
            try { $acroApp.Exit() } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($acroApp) | Out-Null
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function to create basic comparison report
function Create-BasicComparisonReport {
    param(
        [string]$PDF1,
        [string]$PDF2,
        [string]$OutputPath
    )
    
    $reportContent = @"
PDF Comparison Report
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

File 1: $PDF1
File 2: $PDF2

Basic File Information:
- File 1 Size: $((Get-Item $PDF1).Length) bytes
- File 2 Size: $((Get-Item $PDF2).Length) bytes
- File 1 Modified: $((Get-Item $PDF1).LastWriteTime)
- File 2 Modified: $((Get-Item $PDF2).LastWriteTime)

Note: Advanced comparison requires Adobe Acrobat Pro with comparison features.
For detailed comparison, please use Adobe Acrobat's built-in comparison tool manually.
"@
    
    $txtReportPath = $OutputPath.Replace('.pdf', '.txt')
    $reportContent | Out-File -FilePath $txtReportPath -Encoding UTF8
    Write-Host "Basic report saved to: $txtReportPath" -ForegroundColor Green
    
    return $txtReportPath
}

# Main execution
try {
    Write-Host "=== PDF Comparison Script ===" -ForegroundColor Cyan
    Write-Host "Starting PDF comparison process..." -ForegroundColor White
    
    # Validate inputs
    Write-Host "Validating inputs..." -ForegroundColor Yellow
    Test-PDFFile -FilePath $FirstPDF
    Test-PDFFile -FilePath $SecondPDF
    New-OutputDirectory -Path $OutputDirectory
    
    # Check Acrobat installation
    Write-Host "Checking Adobe Acrobat installation..." -ForegroundColor Yellow
    if (-not (Test-AcrobatInstallation)) {
        throw "Adobe Acrobat is not installed or not accessible via COM interface"
    }
    
    # Perform comparison
    $reportPath = Compare-PDFsWithAcrobat -PDF1 $FirstPDF -PDF2 $SecondPDF -OutputDir $OutputDirectory -ReportFileName $ReportName
    
    # Generate summary
    Write-Host "`n=== Comparison Complete ===" -ForegroundColor Green
    Write-Host "First PDF: $FirstPDF" -ForegroundColor White
    Write-Host "Second PDF: $SecondPDF" -ForegroundColor White
    Write-Host "Report Location: $reportPath" -ForegroundColor White
    Write-Host "Output Directory: $OutputDirectory" -ForegroundColor White
    
    # Open output directory
    $openDir = Read-Host "`nWould you like to open the output directory? (y/n)"
    if ($openDir -eq 'y' -or $openDir -eq 'Y') {
        Start-Process explorer.exe -ArgumentList $OutputDirectory
    }
    
    return $reportPath
}
catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    exit 1
}
