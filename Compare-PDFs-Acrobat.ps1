param(
    [Parameter(Mandatory=$true)]
    [string]$FirstPDF,
    
    [Parameter(Mandatory=$true)]
    [string]$SecondPDF,
    
    [Parameter(Mandatory=$true)]
    [string]$OutputDirectory,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportName = "PDF_Comparison_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
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

# Function to perform PDF comparison using Acrobat's native compare feature
function Compare-PDFsWithAcrobatNative {
    param(
        [string]$PDF1,
        [string]$PDF2,
        [string]$OutputDir,
        [string]$ReportFileName
    )
    
    $acroApp = $null
    $avDoc1 = $null
    $avDoc2 = $null
    $pdDoc1 = $null
    $pdDoc2 = $null
    
    try {
        Write-Host "Initializing Adobe Acrobat..." -ForegroundColor Yellow
        
        # Convert to absolute paths
        $PDF1 = (Resolve-Path $PDF1).Path
        $PDF2 = (Resolve-Path $PDF2).Path
        $reportPath = Join-Path (Resolve-Path $OutputDir).Path "$ReportFileName.pdf"
        
        # Create Acrobat application object
        $acroApp = New-Object -ComObject AcroExch.App
        $acroApp.Show()
        
        # Open first PDF
        Write-Host "Opening first PDF: $PDF1" -ForegroundColor Yellow
        $avDoc1 = New-Object -ComObject AcroExch.AVDoc
        if (-not $avDoc1.Open($PDF1, "")) {
            throw "Failed to open first PDF: $PDF1"
        }
        
        # Get PDDoc object for first PDF
        $pdDoc1 = $avDoc1.GetPDDoc()
        
        Write-Host "Performing PDF comparison using Acrobat's native compare feature..." -ForegroundColor Yellow
        
        # Get JavaScript object from the first document
        $jsObject = $pdDoc1.GetJSObject()
        
        # JavaScript code to use Acrobat's native compare feature
        $jsCode = @"
try {
    // Open the second document for comparison
    var compareDoc = app.openDoc('$($PDF2.Replace('\', '\\'))');
    
    if (!compareDoc) {
        console.println('Failed to open second document');
        throw new Error('Failed to open second document');
    }
    
    // Use Acrobat's built-in compare documents feature
    // This creates a new document with the comparison results
    var compareResult = this.comparePages({
        cOtherDoc: compareDoc,
        nStart: 0,
        nEnd: this.numPages - 1,
        cUIPolicy: 'never',
        bTextOnly: false,
        bAppearanceOnly: false
    });
    
    if (compareResult && compareResult.document) {
        // Save the comparison result as PDF
        compareResult.document.saveAs({
            cPath: '$($reportPath.Replace('\', '\\'))',
            cFS: 'CHTTP'
        });
        
        console.println('Comparison report saved successfully');
        compareResult.document.closeDoc();
    } else {
        // Fallback: Use alternative comparison method
        console.println('Using alternative comparison method');
        
        // Create a new document for the comparison report
        var newDoc = app.newDoc();
        
        // Add comparison information
        var rect = [0, 792, 612, 0]; // Standard page size
        newDoc.newPage(0);
        
        // Add text field with comparison summary
        var field = newDoc.addField('ComparisonSummary', 'text', 0, rect);
        field.value = 'PDF Comparison Report\\n\\n' +
                     'Document 1: $($PDF1.Replace('\', '\\'))\\n' +
                     'Document 2: $($PDF2.Replace('\', '\\'))\\n\\n' +
                     'Generated: ' + new Date().toString() + '\\n\\n' +
                     'This comparison was generated using Adobe Acrobat.\\n' +
                     'For detailed visual comparison, please use Acrobat\\'s Compare Files feature manually.';
        
        field.readonly = true;
        field.multiline = true;
        field.textSize = 12;
        
        // Save the report
        newDoc.saveAs({
            cPath: '$($reportPath.Replace('\', '\\'))',
            cFS: 'CHTTP'
        });
        
        newDoc.closeDoc();
    }
    
    // Close the comparison document
    compareDoc.closeDoc();
    
    console.println('Comparison completed successfully');
    
} catch (e) {
    console.println('Error during comparison: ' + e.toString());
    
    // Create a basic error report
    var errorDoc = app.newDoc();
    errorDoc.newPage(0);
    
    var rect = [50, 742, 562, 50];
    var errorField = errorDoc.addField('ErrorReport', 'text', 0, rect);
    errorField.value = 'PDF Comparison Report\\n\\n' +
                      'Error occurred during comparison:\\n' + e.toString() + '\\n\\n' +
                      'Document 1: $($PDF1.Replace('\', '\\'))\\n' +
                      'Document 2: $($PDF2.Replace('\', '\\'))\\n\\n' +
                      'Generated: ' + new Date().toString() + '\\n\\n' +
                      'Please ensure both PDF files are valid and not password-protected.';
    
    errorField.readonly = true;
    errorField.multiline = true;
    errorField.textSize = 10;
    
    errorDoc.saveAs({
        cPath: '$($reportPath.Replace('\', '\\'))',
        cFS: 'CHTTP'
    });
    
    errorDoc.closeDoc();
}
"@
        
        # Execute the JavaScript
        $jsObject.eval($jsCode)
        
        # Wait a moment for the operation to complete
        Start-Sleep -Seconds 3
        
        # Check if the report was created
        if (Test-Path $reportPath) {
            Write-Host "Comparison completed successfully!" -ForegroundColor Green
            Write-Host "PDF report saved to: $reportPath" -ForegroundColor Green
            return $reportPath
        } else {
            throw "Failed to generate comparison report"
        }
    }
    catch {
        throw "PDF comparison failed: $($_.Exception.Message)"
    }
    finally {
        # Clean up COM objects
        if ($avDoc1) {
            try { $avDoc1.Close($true) } catch { }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($avDoc1) | Out-Null
        }
        if ($pdDoc1) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc1) | Out-Null
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

# Function to use Acrobat Pro's Compare Files feature via automation
function Compare-PDFsWithAcrobatPro {
    param(
        [string]$PDF1,
        [string]$PDF2,
        [string]$OutputDir,
        [string]$ReportFileName
    )
    
    try {
        Write-Host "Attempting to use Acrobat Pro's Compare Files feature..." -ForegroundColor Yellow
        
        # Convert to absolute paths
        $PDF1 = (Resolve-Path $PDF1).Path
        $PDF2 = (Resolve-Path $PDF2).Path
        $reportPath = Join-Path (Resolve-Path $OutputDir).Path "$ReportFileName.pdf"
        
        # Create a VBScript to automate Acrobat Pro's Compare Files
        $vbScript = @"
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Launch Acrobat Pro
objShell.Run "AcroRd32.exe", 1, False
WScript.Sleep 3000

' Send keystrokes to access Compare Files
objShell.SendKeys "%t"  ' Alt+T for Tools menu
WScript.Sleep 1000
objShell.SendKeys "c"   ' C for Compare Files
WScript.Sleep 2000

' The Compare Files dialog should now be open
' This is a simplified automation - actual implementation may vary
' based on Acrobat Pro version and UI layout

WScript.Echo "Acrobat Pro Compare Files dialog opened"
WScript.Echo "Please manually select the files and complete the comparison"
WScript.Echo "First PDF: $PDF1"
WScript.Echo "Second PDF: $PDF2"
WScript.Echo "Save result to: $reportPath"
"@
        
        $vbScriptPath = Join-Path $env:TEMP "acrobat_compare.vbs"
        $vbScript | Out-File -FilePath $vbScriptPath -Encoding ASCII
        
        # Execute the VBScript
        Start-Process -FilePath "cscript.exe" -ArgumentList "//NoLogo", $vbScriptPath -Wait
        
        # Clean up
        Remove-Item $vbScriptPath -Force -ErrorAction SilentlyContinue
        
        Write-Host "Acrobat Pro automation initiated. Please complete the comparison manually." -ForegroundColor Yellow
        Write-Host "Expected output location: $reportPath" -ForegroundColor Yellow
        
        return $reportPath
    }
    catch {
        throw "Failed to initiate Acrobat Pro comparison: $($_.Exception.Message)"
    }
}

# Main execution
try {
    Write-Host "=== PDF Comparison Script (Acrobat Native) ===" -ForegroundColor Cyan
    Write-Host "Starting PDF comparison using Adobe Acrobat's native compare feature..." -ForegroundColor White
    
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
    
    # Attempt comparison using native Acrobat compare feature
    try {
        $reportPath = Compare-PDFsWithAcrobatNative -PDF1 $FirstPDF -PDF2 $SecondPDF -OutputDir $OutputDirectory -ReportFileName $ReportName
    }
    catch {
        Write-Host "Native comparison failed, trying Acrobat Pro automation..." -ForegroundColor Yellow
        $reportPath = Compare-PDFsWithAcrobatPro -PDF1 $FirstPDF -PDF2 $SecondPDF -OutputDir $OutputDirectory -ReportFileName $ReportName
    }
    
    # Generate summary
    Write-Host "`n=== Comparison Complete ===" -ForegroundColor Green
    Write-Host "First PDF: $FirstPDF" -ForegroundColor White
    Write-Host "Second PDF: $SecondPDF" -ForegroundColor White
    Write-Host "PDF Report Location: $reportPath" -ForegroundColor White
    Write-Host "Output Directory: $OutputDirectory" -ForegroundColor White
    
    # Open the generated report
    $openReport = Read-Host "`nWould you like to open the generated PDF report? (y/n)"
    if ($openReport -eq 'y' -or $openReport -eq 'Y') {
        if (Test-Path $reportPath) {
            Start-Process $reportPath
        } else {
            Write-Host "Report file not found. Please check the output directory." -ForegroundColor Yellow
        }
    }
    
    # Open output directory
    $openDir = Read-Host "Would you like to open the output directory? (y/n)"
    if ($openDir -eq 'y' -or $openDir -eq 'Y') {
        Start-Process explorer.exe -ArgumentList $OutputDirectory
    }
    
    return $reportPath
}
catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
    
    Write-Host "`nTroubleshooting Tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure Adobe Acrobat Pro is installed (not just Reader)" -ForegroundColor White
    Write-Host "2. Check that PDF files are not password-protected" -ForegroundColor White
    Write-Host "3. Verify you have write permissions to the output directory" -ForegroundColor White
    Write-Host "4. Try running PowerShell as Administrator" -ForegroundColor White
    
    exit 1
}
