# PDF Rescaler v2.0 - A Comprehensive PDF Processing Tool

PDF Rescaler v2.0 is a versatile Python application built with `tkinter` that provides a user-friendly graphical interface for various PDF manipulation tasks. It includes features for zooming/scaling PDF content, splitting PDFs into multiple parts, merging multiple PDFs, and converting Word documents (.docx) to PDF format.

The application is designed to be responsive and prevent UI freezing by loading heavy libraries (`PyMuPDF` and `docx2pdf`) in a background thread.

## Features

* **PDF Zooming/Scaling**:
    * **Custom Zoom**: Scale PDF content by a specified percentage (e.g., 107% to slightly enlarge content while maintaining page size).
    * **Fabuchi Preset**: Apply a predefined crop that effectively zooms into the central 90% of the page, removing 5% margins on all sides.
* **PDF Splitting**:
    * Split a single PDF into multiple files based on user-defined page ranges.
    * Generates both individual split files and a combined PDF that includes blank pages where necessary to ensure even page counts for double-sided printing.
* **PDF Merging**:
    * Combine multiple selected PDF files into a single PDF document.
    * Automatically inserts blank pages after PDFs with an odd number of pages to maintain correct alignment for printing.
* **Word to PDF Conversion**:
    * Convert `.docx` Word documents to PDF format.
    * Includes multiple fallback methods to improve compatibility and handle various system configurations (e.g., COM cleanup for Windows).
* **Background Library Loading**: Prevents GUI freezing during startup by loading heavy libraries in a separate thread.
* **Intuitive User Interface**: A clean and organized `tkinter` GUI with clear labels and status updates.
* **Real-time Status Log**: Provides detailed feedback on processing tasks, including success and error messages, with colored indicators.

## Requirements

Before running the application, you need to install the necessary Python libraries.

* Python 3.x
* `PyMuPDF` (fitz)
* `docx2pdf`
* `tkinter` (usually included with Python standard library)
* `pythoncom` (for advanced Word to PDF conversion on Windows, part of `pywin32`)

You can install the required libraries using pip:

```bash
pip install PyMuPDF docx2pdf pywin32
```

**Note for `docx2pdf`:**
On Windows, `docx2pdf` relies on Microsoft Word being installed on your system.
On macOS, it requires LibreOffice to be installed.
On Linux, it also requires LibreOffice.

## How to Run

1.  **Save the code**: Save the provided code as `MAIN.py`.
2.  **Install dependencies**: Ensure you have installed all the required libraries as mentioned in the "Requirements" section.
3.  **Run the application**: Open a terminal or command prompt, navigate to the directory where you saved `MAIN.py`, and run:

    ```bash
    python MAIN.py
    ```

## Usage

Upon launching the application, you will see the main window with several sections:

### 1. Select PDF Files

* Click the "📁 Browse PDF Files" button to open a file dialog and select one or more PDF files you wish to process.
* The selected files will be listed in the box below, and the "No files selected" label will update to show the count of selected files.
* Click "🗑️ Clear Selection" to remove all selected files from the list.

### 2. Zoom & Process

This section is for scaling the content of selected PDFs.

* **Custom Zoom**:
    * Select the "Custom Zoom" radio button.
    * Enter a percentage in the "Percentage:" field (e.g., `107` for 107% zoom).
* **Fabuchi Preset**:
    * Select the "🎯 Fabuchi Preset" radio button to apply a standard 90% crop (5% margin removal on all sides).
* Click "🔄 Process Selected PDFs" to apply the chosen zoom/crop to all selected PDFs. You will be prompted to choose an output folder.

### 3. PDF Operations

This section contains tools for splitting, merging, and converting files.

* **✂️ Split PDF**:
    * Select **only one** PDF file in the "Select PDF Files" section.
    * Click "✂️ Split PDF".
    * A dialog will appear asking for page ranges (e.g., `1 10 11 20` to split into two parts: pages 1-10 and 11-20).
    * The application will create individual PDF files for each range and a combined PDF (with blank pages inserted for even printing if necessary) in a new subfolder named after the original PDF with `_split` suffix.
* **📎 Merge PDFs**:
    * Select multiple PDF files in the "Select PDF Files" section.
    * Click "📎 Merge PDFs".
    * You will be asked to choose a filename and location for the merged PDF.
    * The tool automatically adds blank pages to ensure each original PDF section starts on a right-hand page when printed.
* **📄 Word → PDF**:
    * Click "📄 Word → PDF".
    * Select one or more `.docx` Word documents.
    * Choose an output folder where the converted PDF files will be saved.
    * The application will attempt to convert each `.docx` file to a PDF.

### 4. Status & Progress

* This text area displays real-time messages about the application's operations, including library loading status, processing progress, and any errors.
* Messages are color-coded: green for success, red for errors, and blue for information.
* Click "Clear Log" to clear the displayed status messages.

## Troubleshooting

* **"Libraries are still loading"**: If you see this message, please wait a few moments. The application loads `PyMuPDF` and `docx2pdf` in the background to prevent the UI from freezing.
* **Word to PDF Conversion Errors**:
    * Ensure Microsoft Word (Windows) or LibreOffice (macOS/Linux) is installed on your system.
    * Try closing all instances of Microsoft Word before attempting conversion.
    * On Windows, try running the application as an administrator.
    * If you encounter `pythoncom` errors, ensure `pywin32` is correctly installed.

## Author

This tool was developed by Lorenzo Liverani (and a lot of different AI tools).

## License

This project is open-source and available to everyone
