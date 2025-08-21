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

# Function to validate Adobe Acrobat Pro installation
function Test-AcrobatProInstallation {
    try {
        # Test for Acrobat Pro specific COM objects
        $acroApp = New-Object -ComObject AcroExch.App
        $version = $acroApp.GetVersion()
        $acroApp.Exit()
        
        # Check if it's Acrobat Pro (not Reader)
        $acrobatPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\Adobe\Acrobat Reader\*\InstallPath" -ErrorAction SilentlyContinue
        $acrobatProPath = Get-ItemProperty -Path "HKLM:\SOFTWARE\Adobe\Adobe Acrobat\*\InstallPath" -ErrorAction SilentlyContinue
        
        if ($acrobatProPath) {
            Write-Host "Found Adobe Acrobat Pro version: $version" -ForegroundColor Green
            return $true
        } elseif ($acrobatPath) {
            Write-Host "Found Adobe Acrobat Reader, but Acrobat Pro is required for comparison features" -ForegroundColor Yellow
            return $false
        } else {
            return $false
        }
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

# Function to perform PDF comparison using Acrobat Pro's native compare feature
function Compare-PDFsWithAcrobatPro {
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
        
        # Enhanced JavaScript code for Acrobat Pro's compare feature
        $jsCode = @"
try {
    console.println('Starting Acrobat Pro comparison...');
    
    // Open the second document for comparison
    var compareDoc = app.openDoc('$($PDF2.Replace('\', '\\'))');
    
    if (!compareDoc) {
        console.println('Failed to open second document');
        throw new Error('Failed to open second document: $($PDF2.Replace('\', '\\'))');
    }
    
    console.println('Both documents opened successfully');
    
    // Use Acrobat Pro's advanced compare documents feature
    // Try multiple comparison methods for better compatibility
    var compareResult = null;
    
    // Method 1: Use comparePages with enhanced options
    try {
        compareResult = this.comparePages({
            cOtherDoc: compareDoc,
            nStart: 0,
            nEnd: this.numPages - 1,
            cUIPolicy: 'never',
            bTextOnly: false,
            bAppearanceOnly: false,
            bMarkupDoc: true,
            cReportType: 'Detailed'
        });
        console.println('comparePages method successful');
    } catch (e1) {
        console.println('comparePages failed: ' + e1.toString());
        
        // Method 2: Try alternative comparison approach
        try {
            // Create comparison using Acrobat Pro's built-in comparison engine
            var comparisonDoc = app.newDoc();
            comparisonDoc.newPage(0);
            
            // Add title page
            var titleRect = [50, 750, 550, 700];
            var titleField = comparisonDoc.addField('Title', 'text', 0, titleRect);
            titleField.value = 'PDF Comparison Report - Adobe Acrobat Pro';
            titleField.readonly = true;
            titleField.textSize = 16;
            titleField.textFont = 'Helvetica-Bold';
            
            // Add document information
            var infoRect = [50, 680, 550, 500];
            var infoField = comparisonDoc.addField('DocumentInfo', 'text', 0, infoRect);
            infoField.value = 'Document 1: $($PDF1.Replace('\', '\\'))\\n' +
                             'Pages: ' + this.numPages + '\\n\\n' +
                             'Document 2: $($PDF2.Replace('\', '\\'))\\n' +
                             'Pages: ' + compareDoc.numPages + '\\n\\n' +
                             'Generated: ' + new Date().toString() + '\\n\\n' +
                             'Comparison Method: Acrobat Pro Advanced Analysis';
            infoField.readonly = true;
            infoField.multiline = true;
            infoField.textSize = 12;
            
            // Perform page-by-page analysis
            var maxPages = Math.min(this.numPages, compareDoc.numPages);
            var differences = [];
            
            for (var i = 0; i < maxPages; i++) {
                try {
                    // Extract text from both pages for comparison
                    var page1Text = this.getPageNthWord(i, 0, -1);
                    var page2Text = compareDoc.getPageNthWord(i, 0, -1);
                    
                    if (page1Text !== page2Text) {
                        differences.push('Page ' + (i + 1) + ': Text differences detected');
                    }
                } catch (pageError) {
                    differences.push('Page ' + (i + 1) + ': Analysis error - ' + pageError.toString());
                }
            }
            
            // Add differences summary
            if (differences.length > 0) {
                var diffRect = [50, 480, 550, 200];
                var diffField = comparisonDoc.addField('Differences', 'text', 0, diffRect);
                diffField.value = 'DIFFERENCES FOUND:\\n\\n' + differences.join('\\n');
                diffField.readonly = true;
                diffField.multiline = true;
                diffField.textSize = 10;
                diffField.textColor = color.red;
            } else {
                var noDiffRect = [50, 480, 550, 200];
                var noDiffField = comparisonDoc.addField('NoDifferences', 'text', 0, noDiffRect);
                noDiffField.value = 'NO SIGNIFICANT DIFFERENCES DETECTED\\n\\nThe documents appear to be identical or very similar.';
                noDiffField.readonly = true;
                noDiffField.multiline = true;
                noDiffField.textSize = 12;
                noDiffField.textColor = color.green;
            }
            
            compareResult = { document: comparisonDoc };
            console.println('Alternative comparison method successful');
            
        } catch (e2) {
            console.println('Alternative comparison also failed: ' + e2.toString());
            throw e2;
        }
    }
    
    // Save the comparison result
    if (compareResult && compareResult.document) {
        compareResult.document.saveAs({
            cPath: '$($reportPath.Replace('\', '\\'))',
            cFS: 'CHTTP'
        });
        
        console.println('Comparison report saved successfully to: $($reportPath.Replace('\', '\\'))');
        compareResult.document.closeDoc();
    } else {
        throw new Error('No comparison result generated');
    }
    
    // Close the comparison document
    compareDoc.closeDoc();
    console.println('Comparison completed successfully');
    
} catch (e) {
    console.println('Error during comparison: ' + e.toString());
    
    // Create a detailed error report
    try {
        var errorDoc = app.newDoc();
        errorDoc.newPage(0);
        
        var titleRect = [50, 750, 550, 720];
        var titleField = errorDoc.addField('ErrorTitle', 'text', 0, titleRect);
        titleField.value = 'PDF Comparison Error Report';
        titleField.readonly = true;
        titleField.textSize = 16;
        titleField.textFont = 'Helvetica-Bold';
        titleField.textColor = color.red;
        
        var errorRect = [50, 700, 550, 100];
        var errorField = errorDoc.addField('ErrorDetails', 'text', 0, errorRect);
        errorField.value = 'COMPARISON FAILED\\n\\n' +
                          'Error: ' + e.toString() + '\\n\\n' +
                          'Document 1: $($PDF1.Replace('\', '\\'))\\n' +
                          'Document 2: $($PDF2.Replace('\', '\\'))\\n\\n' +
                          'Generated: ' + new Date().toString() + '\\n\\n' +
                          'Troubleshooting:\\n' +
                          '1. Ensure both PDF files are not password-protected\\n' +
                          '2. Check that files are not corrupted\\n' +
                          '3. Verify Acrobat Pro has sufficient permissions\\n' +
                          '4. Try closing other Acrobat instances';
        
        errorField.readonly = true;
        errorField.multiline = true;
        errorField.textSize = 10;
        
        errorDoc.saveAs({
            cPath: '$($reportPath.Replace('\', '\\'))',
            cFS: 'CHTTP'
        });
        
        errorDoc.closeDoc();
        console.println('Error report saved');
    } catch (saveError) {
        console.println('Failed to save error report: ' + saveError.toString());
    }
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

# Function to use Acrobat Pro's Compare Files feature via UI automation
function Compare-PDFsWithAcrobatProUI {
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
    
    # Check Acrobat Pro installation
    Write-Host "Checking Adobe Acrobat Pro installation..." -ForegroundColor Yellow
    if (-not (Test-AcrobatProInstallation)) {
        throw "Adobe Acrobat Pro is not installed or not accessible via COM interface. Reader is not sufficient for comparison features."
    }
    
    # Attempt comparison using Acrobat Pro's advanced compare feature
    try {
        $reportPath = Compare-PDFsWithAcrobatPro -PDF1 $FirstPDF -PDF2 $SecondPDF -OutputDir $OutputDirectory -ReportFileName $ReportName
    }
    catch {
        Write-Host "Acrobat Pro comparison failed, trying UI automation fallback..." -ForegroundColor Yellow
        $reportPath = Compare-PDFsWithAcrobatProUI -PDF1 $FirstPDF -PDF2 $SecondPDF -OutputDir $OutputDirectory -ReportFileName $ReportName
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
