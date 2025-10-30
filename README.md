# PYTHON-Assignment

Small project assignment (from Internshala) that extracts text, email addresses, and phone numbers from multiple PDF resumes and writes results into an Excel file.

Contents
- `ASSIGNMENT` — Python script (no .py extension) that reads PDFs from a folder named `my_folder2`, extracts text with `pdfplumber`, finds emails and phone numbers with regex, and writes results to `output.xlsx` using `openpyxl`.
- Several sample PDF resumes (e.g., `AarushiRohatgi.pdf`, `AkashSharma.pdf`, etc.).
- `output.xlsx` — example output produced by the script.
- `ASSIGNMENT_README.md` — more detailed notes for the script (generated).

Quick start
1. Install dependencies (PowerShell):

```powershell
python -m pip install --upgrade pip
python -m pip install pdfplumber openpyxl
```

2. Place PDFs you want to process into a folder named `my_folder2` at the repository root.

3. Run the script. If your shell doesn't execute the `ASSIGNMENT` file directly, rename it to `assignment.py` and run:

```powershell
python .\assignment.py
```

4. The script will create `output.xlsx` in the repository root containing: File Name, Email Addresses, Phone Numbers, and extracted Text.

Notes and suggestions
- The script currently uses simple regexes and will only detect common email and US 10-digit phone formats.
- If PDFs are scanned images (no selectable text), use OCR (e.g., `pytesseract`).
- Consider renaming `ASSIGNMENT` to `assignment.py` and adding `requirements.txt` and CLI argument parsing for better usability.

If you'd like, I can apply the suggested improvements (rename the script, add `requirements.txt`, and add command-line options). Tell me which you'd prefer.
