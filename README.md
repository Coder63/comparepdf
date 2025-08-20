# PDF Comparison PowerShell Script

A PowerShell script that uses Adobe Acrobat's COM interface to compare two PDF files and generate a detailed comparison report.

## Prerequisites

- **Adobe Acrobat Pro** (required for advanced comparison features)
- **Windows PowerShell 5.1** or **PowerShell 7+**
- **Windows OS** (required for COM interface)

## Features

- ✅ Validates PDF files before comparison
- ✅ Uses Adobe Acrobat's native comparison engine
- ✅ Generates timestamped reports
- ✅ Creates output directories automatically
- ✅ Comprehensive error handling
- ✅ Fallback to basic comparison if advanced features fail
- ✅ COM object cleanup to prevent memory leaks

## Usage

### Basic Syntax
```powershell
.\Compare-PDFs.ps1 -FirstPDF "path\to\first.pdf" -SecondPDF "path\to\second.pdf" -OutputDirectory "path\to\output"
```

### Parameters

| Parameter | Required | Description | Default |
|-----------|----------|-------------|---------|
| `FirstPDF` | Yes | Path to the first PDF file | - |
| `SecondPDF` | Yes | Path to the second PDF file | - |
| `OutputDirectory` | Yes | Directory where the report will be saved | - |
| `ReportName` | No | Custom name for the report file | `PDF_Comparison_Report_YYYYMMDD_HHMMSS` |

### Examples

#### Example 1: Basic Comparison
```powershell
.\Compare-PDFs.ps1 -FirstPDF "C:\Documents\original.pdf" -SecondPDF "C:\Documents\revised.pdf" -OutputDirectory "C:\Reports"
```

#### Example 2: Custom Report Name
```powershell
.\Compare-PDFs.ps1 -FirstPDF ".\version1.pdf" -SecondPDF ".\version2.pdf" -OutputDirectory ".\comparison_results" -ReportName "Contract_Changes_Review"
```

#### Example 3: Using Relative Paths
```powershell
.\Compare-PDFs.ps1 -FirstPDF ".\docs\draft1.pdf" -SecondPDF ".\docs\draft2.pdf" -OutputDirectory ".\output"
```

## Output

The script generates:

1. **PDF Comparison Report** (if Acrobat Pro comparison succeeds)
   - Visual comparison with highlighted differences
   - Side-by-side page comparison
   - Summary of changes

2. **Text Report** (fallback or additional info)
   - File information and metadata
   - Basic comparison statistics
   - Timestamps and file sizes

## Error Handling

The script handles various error scenarios:

- **Missing PDF files**: Validates file existence before processing
- **Invalid file types**: Ensures files have `.pdf` extension
- **Acrobat not installed**: Checks for Adobe Acrobat COM interface
- **Permission issues**: Validates write access to output directory
- **COM object failures**: Proper cleanup to prevent memory leaks

## Troubleshooting

### Common Issues

1. **"Adobe Acrobat is not installed"**
   - Install Adobe Acrobat Pro (Reader may not have COM interface)
   - Ensure Acrobat is properly registered

2. **"Failed to open PDF"**
   - Check if PDF files are not password-protected
   - Ensure files are not corrupted
   - Verify file paths are correct

3. **"Access denied to output directory"**
   - Run PowerShell as Administrator
   - Check folder permissions
   - Ensure output directory is not read-only

### PowerShell Execution Policy

If you get execution policy errors, run:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Advanced Usage

### Batch Processing Multiple PDFs
```powershell
# Example batch script
$pdfPairs = @(
    @{First="doc1_v1.pdf"; Second="doc1_v2.pdf"; Output="comparison1"},
    @{First="doc2_v1.pdf"; Second="doc2_v2.pdf"; Output="comparison2"}
)

foreach ($pair in $pdfPairs) {
    .\Compare-PDFs.ps1 -FirstPDF $pair.First -SecondPDF $pair.Second -OutputDirectory $pair.Output
}
```

### Integration with Other Scripts
```powershell
# Call from another script
$reportPath = .\Compare-PDFs.ps1 -FirstPDF $file1 -SecondPDF $file2 -OutputDirectory $outputDir
if ($reportPath) {
    Write-Host "Comparison successful: $reportPath"
    # Process the report further...
}
```

## Technical Details

- **COM Interface**: Uses `AcroExch.App`, `AcroExch.AVDoc`, and `AcroExch.PDDoc`
- **JavaScript Integration**: Leverages Acrobat's JavaScript API for advanced comparison
- **Memory Management**: Proper COM object disposal and garbage collection
- **Error Recovery**: Graceful fallback mechanisms

## License

This script is provided as-is for educational and productivity purposes.

## Support

For issues or improvements, please check:
1. Adobe Acrobat documentation for COM interface
2. PowerShell documentation for COM object handling
3. Ensure all prerequisites are met
