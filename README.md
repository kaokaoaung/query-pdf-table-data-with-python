📌 Overview

This script extracts tabular data from multi-page PDFs and converts it into a single Excel file.

I created this tool because Excel Power Query struggles with large PDFs containing hundreds or even thousands of pages of tables. Using Python and pdfplumber, the script processes each page, extracts tables, and combines them into a single clean Excel file.

⚡ Features

Extracts tables from all pages of a PDF (tested on 1000+ pages).

Automatically combines tables into one structured Excel sheet.

Handles empty tables gracefully (skips them).

Creates the output directory if it doesn’t exist.

Provides progress messages while processing pages.

🚀 How to Use
1. Requirements

Install the required Python libraries:

pip install pdfplumber, pandas, openpyxl

2. Update File Paths

In the script, update these variables:

pdf_path = r"C:\pdf_query\Agent-List-updated-Jan.pdf"  # Input PDF
output_dir = r"C:\pdf_output"                          # Output folder

3. Run the Script
python pdf_to_excel.py


The extracted tables will be saved as:

C:\pdf_output\Agent_List_Extracted.xlsx

✅ Pros & ⚠️ Cons

✅ Pros

Works on large PDFs (1000+ pages) where Excel Power Query often fails.

Fully automated — just set file paths and run.

Handles multiple tables per page.

Simple and lightweight (few dependencies).

Saves output directly in Excel (.xlsx) format.

⚠️ Cons

Table formatting may not be perfect (depends on PDF structure).

Merged cells, headers, or nested tables can cause misaligned data.

Some pages may not contain extractable tables (skipped automatically).

Not optimized for image-based PDFs (OCR needed with something like pytesseract).

Outputs a raw extract — may need post-cleaning in Excel or Power Query.

🔮 Future Improvements

Add header detection and formatting.

Support for OCR-based extraction for scanned PDFs.

Option to save each table into separate Excel sheets.

CLI arguments (so no need to edit file paths in code).
