# PDF Comparison PowerShell Script (Acrobat Native)

A PowerShell script that uses Adobe Acrobat's native compare feature to generate PDF comparison reports.

## Prerequisites

- **Adobe Acrobat Pro DC** (required - Reader will not work)
- **Windows PowerShell 5.1** or **PowerShell 7+**
- **Windows OS** (required for COM interface)
- **Administrator privileges** (recommended)

## Features

- ✅ Uses Acrobat's native `comparePages()` JavaScript API
- ✅ Generates PDF comparison reports (not text files)
- ✅ Visual comparison with highlighted differences
- ✅ Automatic fallback mechanisms
- ✅ Comprehensive error handling
- ✅ COM object cleanup to prevent memory leaks

## Usage

### Basic Syntax
```powershell
.\Compare-PDFs-Acrobat.ps1 -FirstPDF "path\to\first.pdf" -SecondPDF "path\to\second.pdf" -OutputDirectory "path\to\output"
```

### Parameters

| Parameter | Required | Description | Default |
|-----------|----------|-------------|---------|
| `FirstPDF` | Yes | Path to the first PDF file | - |
| `SecondPDF` | Yes | Path to the second PDF file | - |
| `OutputDirectory` | Yes | Directory where the PDF report will be saved | - |
| `ReportName` | No | Custom name for the report file | `PDF_Comparison_YYYYMMDD_HHMMSS` |

### Examples

#### Example 1: Basic Comparison
```powershell
.\Compare-PDFs-Acrobat.ps1 -FirstPDF "C:\Documents\original.pdf" -SecondPDF "C:\Documents\revised.pdf" -OutputDirectory "C:\Reports"
```

#### Example 2: Custom Report Name
```powershell
.\Compare-PDFs-Acrobat.ps1 -FirstPDF ".\contract_v1.pdf" -SecondPDF ".\contract_v2.pdf" -OutputDirectory ".\comparison_results" -ReportName "Contract_Changes_Review"
```

## How It Works

The script uses multiple approaches to ensure successful comparison:

### 1. Primary Method: Acrobat JavaScript API
- Uses `comparePages()` function from Acrobat's JavaScript API
- Generates visual comparison with highlighted differences
- Creates a new PDF document with comparison results

### 2. Fallback Method: UI Automation
- If JavaScript API fails, attempts to automate Acrobat Pro's UI
- Opens Compare Files dialog automatically
- Provides guidance for manual completion

### 3. Error Recovery
- Creates basic comparison report if all methods fail
- Includes file information and error details
- Always generates a PDF output

## Output

The script generates a **PDF comparison report** containing:

1. **Visual Comparison Pages**
   - Side-by-side page comparison
   - Highlighted differences in red/blue
   - Change annotations and comments

2. **Summary Information**
   - File paths and metadata
   - Comparison timestamp
   - Number of differences found

3. **Error Information** (if applicable)
   - Detailed error messages
   - Troubleshooting suggestions

## Advanced Usage

### Batch Processing
```powershell
$pdfPairs = @(
    @{First="doc1_v1.pdf"; Second="doc1_v2.pdf"; Report="Document1_Changes"},
    @{First="doc2_v1.pdf"; Second="doc2_v2.pdf"; Report="Document2_Changes"}
)

foreach ($pair in $pdfPairs) {
    .\Compare-PDFs-Acrobat.ps1 -FirstPDF $pair.First -SecondPDF $pair.Second -OutputDirectory ".\batch_results" -ReportName $pair.Report
}
```

### Integration with Workflows
```powershell
# Automated comparison in a workflow
$reportPath = .\Compare-PDFs-Acrobat.ps1 -FirstPDF $originalDoc -SecondPDF $revisedDoc -OutputDirectory $outputDir -ReportName "Review_$(Get-Date -Format 'yyyyMMdd')"

if (Test-Path $reportPath) {
    # Email the report or process further
    Send-MailMessage -To "reviewer@company.com" -Subject "Document Comparison Report" -Attachments $reportPath
}
```

## Troubleshooting

### Common Issues

1. **"Adobe Acrobat is not installed"**
   - Install Adobe Acrobat Pro DC (not Reader)
   - Ensure Acrobat is properly registered with Windows
   - Try running as Administrator

2. **"Failed to open PDF"**
   - Check if PDFs are password-protected
   - Verify files are not corrupted
   - Ensure file paths don't contain special characters

3. **"Comparison failed"**
   - Close any open Acrobat instances
   - Check available disk space
   - Verify write permissions to output directory

4. **JavaScript Errors**
   - Enable JavaScript in Acrobat preferences
   - Check Acrobat security settings
   - Update to latest Acrobat version

### PowerShell Execution Policy
```powershell
# If you get execution policy errors
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Acrobat Configuration
1. Open Adobe Acrobat Pro
2. Go to Edit → Preferences
3. Select JavaScript
4. Ensure "Enable Acrobat JavaScript" is checked
5. Select Security (Enhanced)
6. Set "Protected Mode" to "Off" (if needed)

## Technical Details

### JavaScript API Used
- `app.openDoc()` - Opens PDF documents
- `comparePages()` - Performs visual comparison
- `saveAs()` - Saves comparison results
- `newDoc()` - Creates new documents for reports

### COM Objects
- `AcroExch.App` - Acrobat application
- `AcroExch.AVDoc` - Document view
- `AcroExch.PDDoc` - PDF document object

### Error Handling
- Multiple fallback mechanisms
- Graceful COM object cleanup
- Detailed error reporting
- User-friendly troubleshooting tips

## Comparison with Other Methods

| Method | Output Format | Visual Diff | Automation | Accuracy |
|--------|---------------|-------------|------------|----------|
| **This Script** | PDF | ✅ Yes | ✅ Full | ✅ High |
| Text-based tools | Text/HTML | ❌ No | ✅ Full | ⚠️ Medium |
| Manual Acrobat | PDF | ✅ Yes | ❌ Manual | ✅ High |

## License

This script is provided as-is for productivity and automation purposes.

## Support

For issues:
1. Verify Adobe Acrobat Pro installation
2. Check PowerShell execution policy
3. Ensure proper file permissions
4. Review Acrobat JavaScript settings
