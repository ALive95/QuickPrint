# The Ultimate PDF Tool 3000

A Python desktop application for PDF manipulation and Word-to-PDF conversion, built with `tkinter`.

## Features

- **PDF Zooming/Scaling**: scale PDF content by a custom percentage, or apply the Fabuchi preset (5% margin crop on all sides)
- **PDF Splitting**: split a single PDF into multiple files by page range; also produces a combined file with blank pages inserted for even duplex page counts
- **PDF Merging**: combine multiple PDFs into one, with automatic blank page insertion for duplex printing alignment
- **Resize to A4**: fit any PDF page to standard A4 dimensions
- **Word to PDF**: convert `.docx` files to PDF using Microsoft Word (Windows) or LibreOffice (macOS/Linux)
- **Background library loading**: heavy libraries load in a background thread so the UI is immediately responsive

## Requirements

- Python 3.x
- `PyMuPDF`
- `docx2pdf`
- `pywin32` (Windows only, for Word-to-PDF fallback via COM)
- `tkinter` (included in the Python standard library)

```bash
pip install PyMuPDF docx2pdf pywin32
```

**Note**: Word-to-PDF conversion requires Microsoft Word on Windows, or LibreOffice on macOS/Linux.

## Running the app

```bash
python MAIN.py
```

## Building the exe (Windows)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=your_icon.ico MAIN.py
```

The executable will be in the `dist/` folder.

## Usage

### 1. Select Documents

Click **Browse** to select one or more `.pdf` or `.docx` files. All operations work on the current selection. If you select a wrong file type for a given operation, the app will warn you and skip the offending files.

### 2. PDF Operations

| Button | What it does |
|---|---|
| Split PDF | Split one selected PDF by page ranges (e.g. `1 10 11 20`) |
| Merge PDFs | Merge all selected PDFs into one file |
| Resize to A4 | Rescale all selected PDFs to A4 page size |

### 3. Zoom & Process

Select a mode and click **Process PDFs**:

- **Custom zoom**: enter a percentage (e.g. `107` for 107%)
- **Fabuchi preset**: crops 5% margins on all sides, effectively zooming into the central 90% of each page

### 4. Word Operations

| Button | What it does |
|---|---|
| Word to PDF | Convert all selected `.docx` files to PDF |

### 5. Clear Selection

Clears the current file selection.

### 6. Log

Real-time status messages for all operations. Green = success, red = error, blue = info. Click **Clear Log** to reset.

## Troubleshooting

- **"Libraries are still loading"**: wait a moment after launch before processing files
- **Word to PDF fails**: make sure Word (Windows) or LibreOffice (macOS/Linux) is installed; try closing all Word instances and/or running as administrator

## Author

Lorenzo Liverani (with a lot of help from AI tools).

## License

Open-source, free to use.
