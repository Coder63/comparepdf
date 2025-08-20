#!/bin/bash

# PDF Comparison Script for macOS
# Uses built-in tools and open-source utilities for PDF comparison

set -e

# Function to display usage
show_usage() {
    echo "Usage: $0 <first_pdf> <second_pdf> <output_directory> [report_name]"
    echo ""
    echo "Parameters:"
    echo "  first_pdf       Path to the first PDF file"
    echo "  second_pdf      Path to the second PDF file"
    echo "  output_directory Directory where the report will be saved"
    echo "  report_name     Optional custom name for the report (default: PDF_Comparison_Report_TIMESTAMP)"
    echo ""
    echo "Examples:"
    echo "  $0 ./one/file1.pdf ./two/file1.pdf ./output"
    echo "  $0 doc1.pdf doc2.pdf ./reports Contract_Changes"
}

# Function to validate PDF file
validate_pdf() {
    local file="$1"
    
    if [[ ! -f "$file" ]]; then
        echo "Error: PDF file not found: $file" >&2
        exit 1
    fi
    
    if [[ ! "$file" =~ \.pdf$ ]]; then
        echo "Error: File is not a PDF: $file" >&2
        exit 1
    fi
    
    # Check if file is a valid PDF using file command
    if ! file "$file" | grep -q "PDF"; then
        echo "Error: Invalid PDF file: $file" >&2
        exit 1
    fi
}

# Function to create output directory
create_output_dir() {
    local dir="$1"
    
    if [[ ! -d "$dir" ]]; then
        mkdir -p "$dir"
        echo "Created output directory: $dir"
    fi
}

# Function to extract text from PDF
extract_pdf_text() {
    local pdf_file="$1"
    local output_file="$2"
    
    # Try different methods to extract text
    if command -v pdftotext >/dev/null 2>&1; then
        pdftotext "$pdf_file" "$output_file"
    elif command -v textutil >/dev/null 2>&1; then
        # macOS built-in textutil
        textutil -convert txt "$pdf_file" -output "$output_file"
    else
        echo "Warning: No text extraction tool found. Install poppler-utils for better results."
        echo "Extracted text not available" > "$output_file"
    fi
}

# Function to get PDF metadata
get_pdf_info() {
    local pdf_file="$1"
    
    echo "File: $pdf_file"
    echo "Size: $(stat -f%z "$pdf_file" 2>/dev/null || stat -c%s "$pdf_file" 2>/dev/null) bytes"
    echo "Modified: $(stat -f%Sm "$pdf_file" 2>/dev/null || stat -c%y "$pdf_file" 2>/dev/null)"
    
    # Try to get PDF info if pdfinfo is available
    if command -v pdfinfo >/dev/null 2>&1; then
        echo "PDF Info:"
        pdfinfo "$pdf_file" 2>/dev/null | head -10
    fi
    
    echo ""
}

# Function to compare PDF files
compare_pdfs() {
    local pdf1="$1"
    local pdf2="$2"
    local output_dir="$3"
    local report_name="$4"
    
    local timestamp=$(date +"%Y%m%d_%H%M%S")
    local report_file="$output_dir/${report_name}_${timestamp}.txt"
    local text1_file="$output_dir/temp_text1.txt"
    local text2_file="$output_dir/temp_text2.txt"
    local diff_file="$output_dir/${report_name}_diff_${timestamp}.txt"
    
    echo "=== PDF Comparison Report ===" > "$report_file"
    echo "Generated: $(date)" >> "$report_file"
    echo "" >> "$report_file"
    
    # Add file information
    echo "=== FILE INFORMATION ===" >> "$report_file"
    get_pdf_info "$pdf1" >> "$report_file"
    get_pdf_info "$pdf2" >> "$report_file"
    
    # Extract text from both PDFs
    echo "Extracting text from PDFs..."
    extract_pdf_text "$pdf1" "$text1_file"
    extract_pdf_text "$pdf2" "$text2_file"
    
    # Compare file sizes
    local size1=$(stat -f%z "$pdf1" 2>/dev/null || stat -c%s "$pdf1" 2>/dev/null)
    local size2=$(stat -f%z "$pdf2" 2>/dev/null || stat -c%s "$pdf2" 2>/dev/null)
    
    echo "=== SIZE COMPARISON ===" >> "$report_file"
    echo "File 1 size: $size1 bytes" >> "$report_file"
    echo "File 2 size: $size2 bytes" >> "$report_file"
    
    if [[ $size1 -eq $size2 ]]; then
        echo "Status: Files are the same size" >> "$report_file"
    else
        local diff=$((size2 - size1))
        echo "Status: File 2 is $diff bytes different from File 1" >> "$report_file"
    fi
    echo "" >> "$report_file"
    
    # Binary comparison
    echo "=== BINARY COMPARISON ===" >> "$report_file"
    if cmp -s "$pdf1" "$pdf2"; then
        echo "Status: Files are identical (binary comparison)" >> "$report_file"
    else
        echo "Status: Files are different (binary comparison)" >> "$report_file"
        
        # Show first few differences
        echo "First differences:" >> "$report_file"
        cmp "$pdf1" "$pdf2" 2>&1 | head -5 >> "$report_file" || true
    fi
    echo "" >> "$report_file"
    
    # Text comparison
    echo "=== TEXT COMPARISON ===" >> "$report_file"
    if [[ -s "$text1_file" && -s "$text2_file" ]]; then
        if diff -q "$text1_file" "$text2_file" >/dev/null 2>&1; then
            echo "Status: Extracted text is identical" >> "$report_file"
        else
            echo "Status: Extracted text differs" >> "$report_file"
            echo "Detailed text differences saved to: $diff_file" >> "$report_file"
            
            # Generate detailed diff
            echo "=== DETAILED TEXT DIFFERENCES ===" > "$diff_file"
            echo "Generated: $(date)" >> "$diff_file"
            echo "" >> "$diff_file"
            
            # Side-by-side diff if available
            if command -v diff >/dev/null 2>&1; then
                echo "--- Unified Diff ---" >> "$diff_file"
                diff -u "$text1_file" "$text2_file" >> "$diff_file" 2>/dev/null || true
                echo "" >> "$diff_file"
                
                echo "--- Side-by-side Diff (first 50 lines) ---" >> "$diff_file"
                diff -y "$text1_file" "$text2_file" | head -50 >> "$diff_file" 2>/dev/null || true
            fi
        fi
    else
        echo "Status: Could not extract text for comparison" >> "$report_file"
    fi
    echo "" >> "$report_file"
    
    # MD5 checksums
    echo "=== CHECKSUMS ===" >> "$report_file"
    local md5_1=$(md5 -q "$pdf1" 2>/dev/null || md5sum "$pdf1" 2>/dev/null | cut -d' ' -f1)
    local md5_2=$(md5 -q "$pdf2" 2>/dev/null || md5sum "$pdf2" 2>/dev/null | cut -d' ' -f1)
    
    echo "File 1 MD5: $md5_1" >> "$report_file"
    echo "File 2 MD5: $md5_2" >> "$report_file"
    
    if [[ "$md5_1" == "$md5_2" ]]; then
        echo "Status: MD5 checksums match (files are identical)" >> "$report_file"
    else
        echo "Status: MD5 checksums differ (files are different)" >> "$report_file"
    fi
    echo "" >> "$report_file"
    
    # Summary
    echo "=== SUMMARY ===" >> "$report_file"
    if [[ "$md5_1" == "$md5_2" ]]; then
        echo "RESULT: The PDF files are identical" >> "$report_file"
    else
        echo "RESULT: The PDF files are different" >> "$report_file"
        echo "- Binary comparison: Different" >> "$report_file"
        echo "- Size difference: $((size2 - size1)) bytes" >> "$report_file"
        if [[ -s "$text1_file" && -s "$text2_file" ]]; then
            if diff -q "$text1_file" "$text2_file" >/dev/null 2>&1; then
                echo "- Text content: Identical" >> "$report_file"
            else
                echo "- Text content: Different" >> "$report_file"
            fi
        fi
    fi
    
    # Clean up temporary files
    rm -f "$text1_file" "$text2_file"
    
    echo "Report saved to: $report_file"
    if [[ -f "$diff_file" ]]; then
        echo "Detailed differences saved to: $diff_file"
    fi
    
    return 0
}

# Main script execution
main() {
    # Check arguments
    if [[ $# -lt 3 ]]; then
        show_usage
        exit 1
    fi
    
    local first_pdf="$1"
    local second_pdf="$2"
    local output_dir="$3"
    local report_name="${4:-PDF_Comparison_Report}"
    
    echo "=== PDF Comparison Script for macOS ==="
    echo "Starting PDF comparison process..."
    
    # Validate inputs
    echo "Validating inputs..."
    validate_pdf "$first_pdf"
    validate_pdf "$second_pdf"
    create_output_dir "$output_dir"
    
    # Perform comparison
    echo "Comparing PDFs..."
    compare_pdfs "$first_pdf" "$second_pdf" "$output_dir" "$report_name"
    
    echo ""
    echo "=== Comparison Complete ==="
    echo "First PDF: $first_pdf"
    echo "Second PDF: $second_pdf"
    echo "Output Directory: $output_dir"
    
    # Offer to open output directory
    if command -v open >/dev/null 2>&1; then
        echo ""
        read -p "Would you like to open the output directory? (y/n): " -n 1 -r
        echo
        if [[ $REPLY =~ ^[Yy]$ ]]; then
            open "$output_dir"
        fi
    fi
}

# Run main function with all arguments
main "$@"
