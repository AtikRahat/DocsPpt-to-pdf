# Docs/PPT to PDF Converter

A Windows desktop app (CustomTkinter) to convert:
- Word files: `.docx`, `.doc`
- PowerPoint files: `.pptx`, `.ppt`

into PDF using Microsoft Office automation.

## Features

- Convert multiple files in one run
- Optional start/end page range
- Custom output filename per file
- Open generated PDF directly from the app
- Friendly GUI (no terminal required)

## Requirements

- Windows 10/11
- Python 3.9+
- Microsoft Word (for `.doc/.docx` conversion)
- Microsoft PowerPoint (for `.ppt/.pptx` conversion)

> This app uses COM automation (`pywin32`), so Office desktop apps must be installed.

## Installation (source)

1. Clone the repository:
   ```bash
   git clone https://github.com/AtikRahat/DocsPpt-to-pdf.git
   cd DocsPpt-to-pdf
   ```

2. (Recommended) Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Run the app:
   ```bash
   python main.py
   ```

## Use the built executable

If you already built the app, run:
- `dist/DocxPptToPdfConverter.exe` (Docx + PPT converter)

## How to use

1. Click **Add Files to Convert**.
2. Select one or more supported files.
3. (Optional) Set **Start Pg** and **End Pg**.
4. (Optional) Change output file name.
5. Click **Convert All**.
6. Click **Open** after conversion to open the generated PDF.

Output PDFs are saved in the same folder as each source file.

## Build executable (PyInstaller)

You can build with the included spec files, for example:
```bash
pyinstaller DocxPptToPdfConverter.spec
```

The executable will appear in `dist/`.

## Troubleshooting

- **"Unsupported file format"**: verify file extension is `.doc/.docx/.ppt/.pptx`.
- **COM/ActiveX errors**: make sure Microsoft Office desktop apps are installed and can open files normally.
- **Permission/path errors**: ensure source files are not read-only and output folder is writable.
- **Range conversion issues**: leave Start/End empty to convert the full document.

## Project structure

- `main.py` – GUI app
- `converter.py` – Word/PowerPoint to PDF conversion logic
- `requirements.txt` – Python dependencies
- `DocxPptToPdfConverter.spec` – PyInstaller build spec

## License

No license file is currently included. Add one if you plan to distribute this project publicly.
