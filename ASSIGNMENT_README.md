# ASSIGNMENT script

This repository contains a Python script named `ASSIGNMENT` (a plain text file containing Python code) which extracts text, email addresses and phone numbers from PDF files and writes the results to an Excel file `output.xlsx`.

## What the script does

- Scans a folder named `my_folder2` inside the repository for PDF files.
- Extracts text from each PDF using `pdfplumber`.
- Finds email addresses and US-format phone numbers using regular expressions.
- Writes the file name, detected email(s), detected phone number(s) and the full extracted text into an Excel workbook (`output.xlsx`) using `openpyxl`.

## Dependencies

The script uses the following Python packages:

- pdfplumber
- openpyxl

Install them with pip before running:

```powershell
python -m pip install --upgrade pip
python -m pip install pdfplumber openpyxl
```

## How to run (Windows PowerShell)

1. Place the PDF files you want to analyze in a folder named `my_folder2` at the repository root (the script expects `my_folder2` in the current working directory).

2. Run the script with Python from the repository root:

```powershell
python .\ASSIGNMENT
```

Note: the script file has no `.py` extension; invoking it as shown works if your environment recognizes it. If that does not work, rename it to `assignment.py` and run:

```powershell
python .\assignment.py
```

3. After successful execution, an Excel file named `output.xlsx` will be created in the repository root containing the extracted data.

## Output format

- Column A: File Name
- Column B: Email Addresses (one per row)
- Column C: Phone Numbers (one per row)
- Column D: Extracted Text (merged across rows for each file when multiple emails/phones are present)

## Notes and caveats

- The email and phone extraction uses simple regular expressions and may miss or mis-detect complex formats.
- The phone regex looks for 10-digit US-style numbers (e.g. `123-456-7890`).
- If PDFs are scanned images without selectable text, `pdfplumber` may not extract text — consider using OCR (e.g., `pytesseract`) for such PDFs.
- The script currently increments rows per email/phone and merges the text column; if you modify it, check the Excel writing logic to avoid overwriting rows.

## Suggestions for improvement (optional)

- Add a `requirements.txt` or rename the script to `assignment.py` for clarity.
- Add argument parsing so the input folder and output path are configurable.
- Add unit tests for the regex extraction functions.

If you'd like, I can:
- Rename the script to `assignment.py` and add a `requirements.txt`.
- Add a small wrapper to accept command-line arguments for input/output paths.

Tell me which of those you'd prefer and I'll implement it.