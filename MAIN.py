# PDF Rescaler v2.0 - A comprehensive PDF processing tool
# Features: PDF zooming, splitting, merging, and Word-to-PDF conversion

import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import threading

# =============================================================================
# GLOBAL VARIABLES AND LIBRARY LOADING
# =============================================================================

# Global variables to track if heavy libraries are loaded
# This prevents UI freezing during startup by loading libraries in background
libraries_loaded = False
fitz = None  # Will hold PyMuPDF library reference
convert_func = None  # Will hold docx2pdf convert function reference


def preload_libraries():
    """
    Preload heavy libraries (PyMuPDF and docx2pdf) in background thread.
    This prevents the GUI from freezing during startup while libraries load.
    Updates the status display once loading is complete or if errors occur.
    """
    global libraries_loaded, fitz, convert_func

    try:
        # Import heavy libraries - this can take several seconds
        import fitz as pymupdf_lib  # PyMuPDF for PDF manipulation
        from docx2pdf import convert as docx_convert  # Word to PDF conversion

        # Store references globally so other functions can use them
        fitz = pymupdf_lib
        convert_func = docx_convert
        libraries_loaded = True

        # Safely update GUI from background thread using root.after()
        root.after(0, update_library_status, "✅ Libraries loaded - Ready to process!", "green")

    except Exception as e:
        # Handle any import errors and display to user
        root.after(0, update_library_status, f"❌ Error loading libraries: {e}", "red")


def update_library_status(message, color):
    """
    Safely update status display from background thread.
    Uses thread-safe method to update GUI elements.

    Args:
        message (str): Status message to display
        color (str): Color tag for text formatting
    """
    status_text.config(state=tk.NORMAL)  # Enable editing
    status_text.insert(tk.END, f"{message}\n", color)  # Add colored message
    status_text.config(state=tk.DISABLED)  # Disable editing
    status_text.yview(tk.END)  # Scroll to bottom


def check_libraries_loaded():
    """
    Check if required libraries are loaded before attempting PDF operations.
    Prevents crashes if user tries to process files before libraries are ready.

    Returns:
        bool: True if libraries are loaded, False otherwise
    """
    if not libraries_loaded:
        messagebox.showwarning("Please Wait",
                               "Libraries are still loading in the background. Please try again in a moment.")
        return False
    return True


def ask_folder_name(default_name):
    top = tk.Toplevel(root)
    top.title("Choose Output Folder Name")
    top.geometry("400x150")
    top.transient(root)
    top.grab_set()
    top.focus_force()

    # Center the window on the screen
    top.update_idletasks()
    screen_w = top.winfo_screenwidth()
    screen_h = top.winfo_screenheight()
    win_w = 400
    win_h = 150
    x = (screen_w - win_w) // 2
    y = (screen_h - win_h) // 2
    top.geometry(f"{win_w}x{win_h}+{x}+{y}")

    tk.Label(top, text="Enter name for output folder. \n"
                       "The folder will be saved in the same location of the PDF file").pack(pady=10)
    entry = tk.Entry(top, width=40)
    entry.insert(0, default_name)
    entry.pack()

    result = {"name": None}

    def on_ok(event=None):
        result["name"] = entry.get().strip()
        top.destroy()

    tk.Button(top, text="OK", command=on_ok).pack(pady=10)

    entry.bind("<Return>", on_ok)
    entry.focus()

    root.wait_window(top)
    return result["name"]


# =============================================================================
# PDF PROCESSING FUNCTIONS
# =============================================================================

def fabuchi_clip_rect(width, height):
    """
    Define a custom scaling rectangle for the "Fabuchi" preset mode.
    This creates a 90% crop of the original page, removing 5% margins on all sides.

    Args:
        width (float): Original page width
        height (float): Original page height

    Returns:
        fitz.Rect: Rectangle defining the crop area
    """
    return fitz.Rect(
        width * 0.05,  # Left margin (5% from left edge)
        height * 0.05,  # Top margin (5% from top edge)
        width * 0.95,  # Right margin (95% from left edge)
        height * 0.95  # Bottom margin (95% from top edge)
    )


def zoom_pdf_content(input_path, output_folder, scale_factor=None, fabuchi=False):
    """
    Main function to zoom/scale PDF content while maintaining original page dimensions.
    Can either use a custom scale factor or apply the Fabuchi preset.

    Args:
        input_path (str): Path to input PDF file
        output_folder (str): Directory to save processed PDF
        scale_factor (float, optional): Custom zoom factor (e.g., 1.07 for 107%)
        fabuchi (bool): Whether to use Fabuchi preset instead of custom scaling
    """
    # Ensure libraries are loaded before processing
    if not check_libraries_loaded():
        return

    # Create output directory if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)

    # Generate output filename with appropriate suffix
    name, ext = os.path.splitext(os.path.basename(input_path))
    suffix = "_fabuchi" if fabuchi else "_zoomed"
    output_path = os.path.join(output_folder, f"{name}{suffix}{ext}")

    def update_status_text(text, color):
        """Thread-safe status update helper function"""
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, text, color)
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)

    try:
        # Open source PDF and create new PDF with same page sizes
        pdf_document = fitz.open(input_path)
        new_pdf = fitz.open()  # Create empty PDF

        # Process each page individually
        for page_num in range(len(pdf_document)):
            original_page = pdf_document[page_num]

            # Create new page with same dimensions as original
            new_page = new_pdf.new_page(width=original_page.rect.width,
                                        height=original_page.rect.height)

            width, height = original_page.rect.width, original_page.rect.height

            if fabuchi:
                # Use predefined Fabuchi cropping rectangle
                clip_rect = fabuchi_clip_rect(width, height)
            else:
                # Calculate custom zoom clip rectangle
                # This centers the zoomed content within the original page bounds
                clip_rect = fitz.Rect(
                    -width * (scale_factor - 1) / 2,  # Left boundary
                    -height * (scale_factor - 1) / 2,  # Top boundary
                    width * scale_factor - width * (scale_factor - 1) / 2,  # Right boundary
                    height * scale_factor - height * (scale_factor - 1) / 2  # Bottom boundary
                )

            # Copy original page content to new page with specified clipping
            new_page.show_pdf_page(clip_rect, pdf_document, page_num, keep_proportion=True)

        # Save the processed PDF
        new_pdf.save(output_path)
        pdf_document.close()
        new_pdf.close()

        # Update status with success message
        root.after(0, update_status_text, f"Processed: {os.path.basename(output_path)}\n", "green")

    except Exception as e:
        # Handle any processing errors
        root.after(0, update_status_text, f"Error processing {input_path}: {e}\n", "red")


def process_pdfs_in_thread():
    """
    Wrapper function to run PDF processing in a separate thread.
    Prevents GUI freezing during processing and manages button states.
    """
    process_button.config(state=tk.DISABLED)  # Disable button during processing
    try:
        process_pdfs()  # Run the actual processing
    finally:
        process_button.config(state=tk.NORMAL)  # Re-enable button when done


def process_pdfs():
    """
    Main processing function that handles user input validation and coordinates
    the processing of all selected PDF files based on chosen mode and settings.
    """
    # Validate that files are selected
    if not selected_files:
        messagebox.showerror("Error", "No PDFs selected!")
        return

    # Get processing mode (zoom or fabuchi)
    mode = mode_var.get()

    if mode == "zoom":
        # Validate zoom percentage input
        try:
            scale_factor = float(zoom_entry.get()) / 100  # Convert percentage to decimal
            if scale_factor <= 0:
                raise ValueError("Scale factor must be positive")
        except ValueError:
            messagebox.showerror("Error", "Enter a valid zoom percentage!")
            return
    else:
        scale_factor = None  # Fabuchi mode doesn't use custom scale factor

    # Let user choose where to save processed files
    output_folder = filedialog.askdirectory(
        title="Choose folder to save processed PDFs"
    )

    if not output_folder:  # User cancelled the dialog
        return

    # Process each selected PDF file
    for file in selected_files:
        zoom_pdf_content(file, output_folder, scale_factor, fabuchi=(mode == "fabuchi"))

    # Display completion message
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, f"Processing complete! Files saved to: {output_folder}\n", "green")
    status_text.config(state=tk.DISABLED)


# =============================================================================
# FILE SELECTION AND MANAGEMENT
# =============================================================================

def select_pdfs():
    """
    Open file dialog to allow user to select multiple PDF files.
    Updates the file listbox and global selected_files variable.
    """
    global selected_files
    selected_files = filedialog.askopenfilenames(
        title="Select PDFs",
        filetypes=[("PDF Files", "*.pdf")]
    )

    # Update the display listbox with selected filenames
    file_listbox.delete(0, tk.END)  # Clear existing entries
    for f in selected_files:
        file_listbox.insert(tk.END, os.path.basename(f))  # Show only filename, not full path


# =============================================================================
# PDF SPLITTING FUNCTIONALITY
# =============================================================================

def split_pdf():
    """
    Split a single PDF into multiple parts based on user-defined page ranges.
    Creates both individual split files and a combined file with blank pages
    inserted to ensure even page counts for printing.
    """
    if not check_libraries_loaded():
        return

    # Validate selection - only one PDF can be split at a time
    if not selected_files:
        messagebox.showerror("Error", "No PDFs selected!")
        return

    if len(selected_files) > 1:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Error: Select only one PDF for splitting!\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)
        return

    # Open PDF and get page count
    pdf_path = selected_files[0]
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)

    # Get page ranges from user input
    input_ranges = simpledialog.askstring(
        "Page Ranges",
        f"Total pages: {total_pages}\nEnter 2n numbers (space-separated) defining the page ranges:\n"
        f"Example: '1 10 11 20' creates two parts: pages 1-10 and pages 11-20"
    )

    if not input_ranges:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Error: No page ranges provided!\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)
        return

    # Parse and validate input
    try:
        page_numbers = list(map(int, input_ranges.split()))
        if len(page_numbers) % 2 != 0:  # Must be even number of values
            raise ValueError("Must provide pairs of start-end page numbers")
    except ValueError:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Error: Invalid input! Enter an even number of integers.\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)
        return

    # Validate each page range
    for i in range(0, len(page_numbers), 2):
        start, end = page_numbers[i], page_numbers[i + 1]
        if start < 1 or end > total_pages or start > end:
            status_text.config(state=tk.NORMAL)
            status_text.insert(tk.END, f"Error: Invalid range {start}-{end}.\n", "red")
            status_text.config(state=tk.DISABLED)
            status_text.yview(tk.END)
            return

    # Choose output directory
    default_name = os.path.splitext(os.path.basename(pdf_path))[0] + "_split"
    folder_name = ask_folder_name(default_name)
    if not folder_name:
        return

    pdf_dir = os.path.dirname(pdf_path)
    output_folder = os.path.join(pdf_dir, folder_name)
    os.makedirs(output_folder, exist_ok=True)

    # Perform the splitting operation
    individual_outputs, combined_output = split_pdf_with_combined_output(pdf_path, page_numbers, output_folder)
    pdf_document.close()

    # Display success message
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, f"PDF split successfully! Files saved to: {output_folder}\n", "green")
    status_text.config(state=tk.DISABLED)
    status_text.yview(tk.END)


def split_pdf_with_combined_output(pdf_path, page_ranges, output_folder):
    """
    Perform the actual PDF splitting operation.
    Creates individual split files AND a combined file with blank pages
    inserted after sections with odd page counts (for printing purposes).

    Args:
        pdf_path (str): Path to source PDF
        page_ranges (list): List of page numbers defining ranges [start1, end1, start2, end2, ...]
        output_folder (str): Directory to save output files

    Returns:
        tuple: (list of individual output paths, combined output path)
    """
    if not check_libraries_loaded():
        return

    pdf_document = fitz.open(pdf_path)
    base_name, _ = os.path.splitext(os.path.basename(pdf_path))
    individual_outputs = []
    merged_pdf = fitz.open()  # For combined output

    # Process each page range pair
    # Process non-overlapping page range pairs
    for i in range(0, len(page_ranges), 2):
        start, end = page_ranges[i], page_ranges[i + 1]
        new_pdf = fitz.open()

        # Copy the full range in one call
        new_pdf.insert_pdf(pdf_document, from_page=start - 1, to_page=end - 1)

        # Save individual split file
        output_path = os.path.join(output_folder, f"{base_name}_{start}-{end}.pdf")
        new_pdf.save(output_path)
        new_pdf.close()
        individual_outputs.append(output_path)

        # Add to combined PDF
        to_merge = fitz.open(output_path)
        merged_pdf.insert_pdf(to_merge)

        # Add blank page if number of pages in section is odd
        if (end - start + 1) % 2 != 0:
            blank = merged_pdf.new_page(width=to_merge[0].rect.width,
                                        height=to_merge[0].rect.height)
            blank.draw_rect(blank.rect)  # Optional: border for blank page

        to_merge.close()

    # Save combined PDF with blank pages
    merged_output_path = os.path.join(output_folder, f"{base_name}_merged.pdf")
    merged_pdf.save(merged_output_path)
    merged_pdf.close()
    pdf_document.close()

    return individual_outputs, merged_output_path


# =============================================================================
# PDF MERGING FUNCTIONALITY
# =============================================================================

def merge_selected_pdfs():
    """
    Merge all selected PDF files into a single PDF.
    Adds blank pages after PDFs with odd page counts to maintain printing alignment.
    """
    if not check_libraries_loaded():
        return

    if not selected_files:
        messagebox.showerror("Error", "No PDFs selected!")
        return

    # Let user choose output filename and location
    output_path = filedialog.asksaveasfilename(
        title="Save Merged PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        initialfile="merged_output.pdf"
    )

    if not output_path:  # User cancelled
        return

    try:
        merged = fitz.open()  # Create new empty PDF

        for file in selected_files:
            pdf = fitz.open(file)
            merged.insert_pdf(pdf)  # Add all pages from this PDF

            # Add blank page if the inserted PDF has odd number of pages
            # This ensures each PDF section starts on the right-hand page when printed
            if len(pdf) % 2 != 0:
                rect = pdf[0].rect  # Use same dimensions as first page
                blank = merged.new_page(width=rect.width, height=rect.height)
                blank.draw_rect(blank.rect)  # Draw border on blank page

            pdf.close()

        # Save merged result
        merged.save(output_path)
        merged.close()

        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, f"Merged PDFs saved as: {os.path.basename(output_path)}\n", "green")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)

    except Exception as e:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, f"Error merging PDFs: {e}\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)


# =============================================================================
# WORD TO PDF CONVERSION
# =============================================================================

def convert_word_to_pdf():
    """
    Convert selected Word documents (.docx) to PDF format.
    Uses multiple fallback methods to handle different system configurations.
    """
    if not check_libraries_loaded():
        return

    # Select Word files
    word_files = filedialog.askopenfilenames(
        title="Select Word Files",
        filetypes=[("Word Documents", "*.docx")]
    )

    if not word_files:
        return

    # Choose output directory
    output_folder = filedialog.askdirectory(
        title="Choose folder to save converted PDFs"
    )

    if not output_folder:
        return

    # Convert each Word file using multiple fallback methods
    for docx_path in word_files:
        success = False
        error_messages = []

        # Method 1: Try docx2pdf with explicit output path
        try:
            import os
            output_name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            output_path = os.path.join(output_folder, output_name)

            # Use convert_func with explicit output path
            convert_func(docx_path, output_path)

            status_text.config(state=tk.NORMAL)
            status_text.insert(tk.END, f"✅ Converted: {os.path.basename(docx_path)}\n", "green")
            status_text.config(state=tk.DISABLED)
            success = True

        except Exception as e1:
            error_messages.append(f"Method 1 (docx2pdf explicit): {str(e1)}")

            # Method 2: Try docx2pdf with folder only
            try:
                convert_func(docx_path, output_folder)

                status_text.config(state=tk.NORMAL)
                status_text.insert(tk.END, f"✅ Converted: {os.path.basename(docx_path)}\n", "green")
                status_text.config(state=tk.DISABLED)
                success = True

            except Exception as e2:
                error_messages.append(f"Method 2 (docx2pdf folder): {str(e2)}")

                # Method 3: Try with COM cleanup
                try:
                    import pythoncom
                    pythoncom.CoInitialize()

                    try:
                        convert_func(docx_path, output_folder)
                        success = True

                        status_text.config(state=tk.NORMAL)
                        status_text.insert(tk.END, f"✅ Converted: {os.path.basename(docx_path)}\n", "green")
                        status_text.config(state=tk.DISABLED)

                    finally:
                        pythoncom.CoUninitialize()

                except Exception as e3:
                    error_messages.append(f"Method 3 (COM cleanup): {str(e3)}")

        # If all methods failed, show error
        if not success:
            status_text.config(state=tk.NORMAL)
            status_text.insert(tk.END, f"❌ Failed to convert: {os.path.basename(docx_path)}\n", "red")
            status_text.insert(tk.END, f"   Tried {len(error_messages)} methods. Last error: {error_messages[-1]}\n",
                               "red")
            status_text.insert(tk.END,
                               f"   💡 Try: Close Word completely, run as administrator, or install python-docx2txt\n",
                               "blue")
            status_text.config(state=tk.DISABLED)

    status_text.yview(tk.END)


# =============================================================================
# GUI SETUP AND LAYOUT
# =============================================================================

# Create main window with modern styling
root = tk.Tk()
root.title("🔥THE ULTIMATE PDF TOOL 3000💥")
root.geometry("750x950")
root.resizable(True, True)

# Configure grid weights for responsive design
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Modern color scheme
bg_color = "#f8f9fa"
accent_color = "#007bff"
root.configure(bg=bg_color)

# Main container frame
main_frame = ttk.Frame(root, padding=15)
main_frame.pack(fill="both", expand=True)
main_frame.columnconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

# =============================================================================
# FILE SELECTION SECTION
# =============================================================================

file_frame = ttk.LabelFrame(main_frame, text="📂 Select PDF Files", padding=15)
file_frame.grid(row=0, column=0, columnspan=2, pady=(0, 15), sticky="ew")

# Button frame for file selection controls
button_frame = ttk.Frame(file_frame)
button_frame.pack(fill="x", pady=(0, 10))

select_button = ttk.Button(button_frame, text="📁 Browse PDF Files", command=select_pdfs)
select_button.pack(side="left")

# File count display label
file_count_label = ttk.Label(button_frame, text="No files selected", foreground="gray")
file_count_label.pack(side="right")

# File list display with scrollbar
listbox_frame = ttk.Frame(file_frame)
listbox_frame.pack(fill="both", expand=True)

file_listbox = tk.Listbox(listbox_frame, height=6, width=70, selectmode=tk.EXTENDED)
scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=file_listbox.yview)
file_listbox.configure(yscrollcommand=scrollbar.set)

file_listbox.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# =============================================================================
# ZOOM & PROCESS SECTION
# =============================================================================

zoom_frame = ttk.LabelFrame(main_frame, text="🔍 Zoom & Process", padding=15)
zoom_frame.grid(row=1, column=0, padx=(0, 8), pady=(0, 15), sticky="nsew")

# Mode selection radio buttons
mode_var = tk.StringVar(value="zoom")  # Default to custom zoom mode

# Custom zoom controls
custom_frame = ttk.Frame(zoom_frame)
custom_frame.pack(fill="x", pady=(0, 10))

ttk.Radiobutton(custom_frame, text="Custom Zoom", variable=mode_var, value="zoom").pack(anchor="w")

# Zoom percentage input
zoom_input_frame = ttk.Frame(custom_frame)
zoom_input_frame.pack(fill="x", padx=(20, 0), pady=(5, 0))

ttk.Label(zoom_input_frame, text="Percentage:").pack(side="left")
zoom_entry = ttk.Entry(zoom_input_frame, width=8)
zoom_entry.pack(side="left", padx=(5, 0))
zoom_entry.insert(0, "107")  # Default zoom value
ttk.Label(zoom_input_frame, text="%").pack(side="left", padx=(2, 0))

# Fabuchi preset option
ttk.Radiobutton(zoom_frame, text="🎯 Fabuchi Preset", variable=mode_var, value="fabuchi").pack(anchor="w", pady=(10, 0))

# Main process button
process_button = ttk.Button(zoom_frame, text="🔄 Process Selected PDFs", command=process_pdfs_in_thread)
process_button.pack(fill="x", pady=(20, 0))

# =============================================================================
# PDF TOOLS SECTION
# =============================================================================

tools_frame = ttk.LabelFrame(main_frame, text="🛠️ PDF Operations", padding=15)
tools_frame.grid(row=1, column=1, padx=(8, 0), pady=(0, 15), sticky="nsew")

# PDF operation buttons
split_button = ttk.Button(tools_frame, text="✂️ Split PDF", command=split_pdf)
split_button.pack(fill="x", pady=(0, 8))

merge_button = ttk.Button(tools_frame, text="📎 Merge PDFs", command=merge_selected_pdfs)
merge_button.pack(fill="x", pady=(0, 8))

word_convert_button = ttk.Button(tools_frame, text="📄 Word → PDF", command=convert_word_to_pdf)
word_convert_button.pack(fill="x", pady=(0, 8))

# Clear selection utility button
clear_button = ttk.Button(tools_frame, text="🗑️ Clear Selection", command=lambda: clear_selection())
clear_button.pack(fill="x", pady=(8, 0))

# =============================================================================
# STATUS AND PROGRESS SECTION
# =============================================================================

status_frame = ttk.LabelFrame(main_frame, text="📋 Status & Progress", padding=15)
status_frame.grid(row=2, column=0, columnspan=2, pady=(0, 0), sticky="ew")

# Status text area with scrollbar
status_container = ttk.Frame(status_frame)
status_container.pack(fill="both", expand=True)

status_text = tk.Text(status_container, height=8, width=80, state=tk.DISABLED, wrap=tk.WORD,
                      font=("Consolas", 9), bg="#ffffff", relief="sunken", borderwidth=1)
status_scrollbar = ttk.Scrollbar(status_container, orient="vertical", command=status_text.yview)
status_text.configure(yscrollcommand=status_scrollbar.set)

status_text.pack(side="left", fill="both", expand=True)
status_scrollbar.pack(side="right", fill="y")

# Configure text color tags for different message types
status_text.tag_configure("green", foreground="#28a745")  # Success messages
status_text.tag_configure("red", foreground="#dc3545")  # Error messages
status_text.tag_configure("blue", foreground="#007bff")  # Info messages

# Clear status log button
clear_status_button = ttk.Button(status_frame, text="Clear Log", command=lambda: clear_status())
clear_status_button.pack(pady=(10, 0))


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def clear_selection():
    """Clear all selected files and update display"""
    global selected_files
    selected_files = []
    file_listbox.delete(0, tk.END)
    file_count_label.config(text="No files selected")


def clear_status():
    """Clear the status/log display area"""
    status_text.config(state=tk.NORMAL)
    status_text.delete(1.0, tk.END)
    status_text.config(state=tk.DISABLED)


def update_select_pdfs():
    """
    Enhanced file selection function that updates the file count display.
    This replaces the basic select_pdfs function with additional UI feedback.
    """
    global selected_files
    selected_files = filedialog.askopenfilenames(
        title="Select PDF Files",
        filetypes=[("PDF Files", "*.pdf")]
    )

    # Update file list display
    file_listbox.delete(0, tk.END)
    for f in selected_files:
        file_listbox.insert(tk.END, os.path.basename(f))

    # Update file count label with appropriate styling
    count = len(selected_files)
    if count == 0:
        file_count_label.config(text="No files selected", foreground="gray")
    elif count == 1:
        file_count_label.config(text="1 file selected", foreground="green")
    else:
        file_count_label.config(text=f"{count} files selected", foreground="green")


# Initialize global variable for selected files
selected_files = []


def show_startup_message():
    """Display initial startup message while libraries load"""
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, "🔄 Loading libraries in background...\n", "blue")
    status_text.config(state=tk.DISABLED)


# =============================================================================
# APPLICATION STARTUP
# =============================================================================

# Show startup message after brief delay to allow GUI to render
root.after(100, show_startup_message)

# Start library loading in background thread (daemon thread closes with main program)
threading.Thread(target=preload_libraries, daemon=True).start()

# Start the GUI event loop
root.mainloop()