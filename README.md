# PDF Estimate to Excel

Convert estimate PDFs into Excel workbooks with three practical outputs:

- a structured `Estimate` sheet for the tabular data;
- a `Header` sheet that keeps the document header as an embedded image;
- a `Source Text` sheet so the full document text is still preserved as far as the parser/OCR can recover it.

The project is built for Russian estimate PDFs, including mixed inputs where some pages have a text layer and some pages are scans.

## Repository layout

```text
.
├── run.bat
├── project/
│   ├── main.py
│   ├── extractor.py
│   ├── ocr_extractor.py
│   ├── converter.py
│   ├── document_export.py
│   ├── quality_report.py
│   ├── config.py
│   ├── requirements.txt
│   ├── tessdata/
│   └── tests/
```

Runtime folders are created next to `run.bat` when needed:

- `input/`
- `output/`

They are intentionally ignored by Git.

## What the tool produces

For each source PDF the converter writes one `.xlsx` workbook with:

- `Header` — image-based header export from the first PDF page.
- `Estimate` — best-effort structured extraction of estimate rows and values.
- `Source Text` — full extracted text from the PDF text layer or OCR fallback.

It also writes:

- `conversion_report.xlsx`
- `conversion_report.json`

Those reports summarize visible-character preservation for the generated output.

## Quick start

### Windows launcher

1. Put PDF files into `input/`.
2. Double-click `run.bat`.
3. Collect the generated `.xlsx` files from `output/`.

### Command line

From `project/`:

```powershell
venv\Scripts\python.exe main.py --input "..\\input" --output "..\\output"
```

Process a single file:

```powershell
venv\Scripts\python.exe main.py --file "..\\input\example.pdf" --output "..\\output"
```

Force OCR for all files:

```powershell
venv\Scripts\python.exe main.py --input "..\\input" --output "..\\output" --use-ocr
```

## Requirements

- Windows
- Python 3.11+
- Tesseract OCR installed at:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

The repository already includes `project/tessdata/rus.traineddata`.

## Verification

Run the test suite from `project/`:

```powershell
venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"
```

## Practical limitations

This tool is reliable for extracting tabular content and for preserving recovered text. It is not a pixel-perfect PDF-to-Excel layout clone.

In practice:

- table pages convert reasonably well into structured Excel rows;
- title/header pages, stamps, signatures, and complex visual layouts do not map cleanly into native Excel cells, so the header is exported as an image sheet instead;
- scanned pages depend heavily on OCR quality.

That is why the workbook contains a structured sheet plus visual/text fallbacks instead of pretending the full PDF can be rebuilt as a clean native Excel layout.
