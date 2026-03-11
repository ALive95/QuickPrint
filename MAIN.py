# PDF Rescaler v2.0 - PDF zooming, splitting, merging, and Word-to-PDF conversion

import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import threading

# =============================================================================
# GLOBALS
# =============================================================================

libraries_loaded = False
fitz = None
convert_func = None
selected_files = []


# =============================================================================
# LIBRARY LOADING
# =============================================================================

def preload_libraries():
    global libraries_loaded, fitz, convert_func
    try:
        import fitz as pymupdf_lib
        from docx2pdf import convert as docx_convert
        fitz = pymupdf_lib
        convert_func = docx_convert
        libraries_loaded = True
        root.after(0, log_status, "Libraries loaded - Ready to process!\n", "green")
    except Exception as e:
        root.after(0, log_status, f"Error loading libraries: {e}\n", "red")


def check_libraries_loaded():
    if not libraries_loaded:
        messagebox.showwarning("Please Wait", "Libraries are still loading. Please try again in a moment.")
        return False
    return True


# =============================================================================
# STATUS HELPER
# =============================================================================

def log_status(message, color="green"):
    """Thread-safe status log update."""
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, message, color)
    status_text.config(state=tk.DISABLED)
    status_text.yview(tk.END)


# =============================================================================
# FILE SELECTION
# =============================================================================

def select_documents():
    global selected_files
    selected_files = filedialog.askopenfilenames(
        title="Select Documents",
        filetypes=[("All supported", "*.pdf *.docx"), ("PDF Files", "*.pdf"), ("Word Documents", "*.docx")]
    )
    file_listbox.delete(0, tk.END)
    for f in selected_files:
        file_listbox.insert(tk.END, os.path.basename(f))

    count = len(selected_files)
    if count == 0:
        file_count_label.config(text="No files selected", foreground="gray")
    elif count == 1:
        file_count_label.config(text="1 file selected", foreground="green")
    else:
        file_count_label.config(text=f"{count} files selected", foreground="green")


def check_all_pdfs():
    """Warn about non-PDF files. Returns True if at least one PDF is selected."""
    non_pdf = [f for f in selected_files if not f.lower().endswith(".pdf")]
    if non_pdf:
        names = ", ".join(os.path.basename(f) for f in non_pdf)
        messagebox.showwarning("Wrong file type", f"This operation only works on PDF files. Skipping:\n{names}")
    return len(non_pdf) < len(selected_files)


def check_all_docx():
    """Warn about non-Word files. Returns True if at least one .docx is selected."""
    non_docx = [f for f in selected_files if not f.lower().endswith(".docx")]
    if non_docx:
        names = ", ".join(os.path.basename(f) for f in non_docx)
        messagebox.showwarning("Wrong file type", f"This operation only works on Word (.docx) files. Skipping:\n{names}")
    return len(non_docx) < len(selected_files)


def clear_selection():
    global selected_files
    selected_files = []
    file_listbox.delete(0, tk.END)
    file_count_label.config(text="No files selected", foreground="gray")


# =============================================================================
# ZOOM / SCALE
# =============================================================================

def fabuchi_clip_rect(width, height):
    """90% crop centered on the page (5% margin on each side)."""
    return fitz.Rect(width * 0.05, height * 0.05, width * 0.95, height * 0.95)


def zoom_pdf_content(input_path, output_folder, scale_factor=None, fabuchi=False):
    if not check_libraries_loaded():
        return

    os.makedirs(output_folder, exist_ok=True)
    name, ext = os.path.splitext(os.path.basename(input_path))
    suffix = "_fabuchi" if fabuchi else "_zoomed"
    output_path = os.path.join(output_folder, f"{name}{suffix}{ext}")

    try:
        pdf_document = fitz.open(input_path)
        new_pdf = fitz.open()

        for page_num in range(len(pdf_document)):
            original_page = pdf_document[page_num]
            width, height = original_page.rect.width, original_page.rect.height
            new_page = new_pdf.new_page(width=width, height=height)

            if fabuchi:
                clip_rect = fabuchi_clip_rect(width, height)
            else:
                # Center zoomed content within original page bounds
                clip_rect = fitz.Rect(
                    -width * (scale_factor - 1) / 2,
                    -height * (scale_factor - 1) / 2,
                    width * scale_factor - width * (scale_factor - 1) / 2,
                    height * scale_factor - height * (scale_factor - 1) / 2
                )

            new_page.show_pdf_page(clip_rect, pdf_document, page_num, keep_proportion=True)

        new_pdf.save(output_path)
        pdf_document.close()
        new_pdf.close()
        root.after(0, log_status, f"Processed: {os.path.basename(output_path)}\n", "green")

    except Exception as e:
        root.after(0, log_status, f"Error processing {input_path}: {e}\n", "red")


def process_pdfs():
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return
    if not check_all_pdfs():
        return

    mode = mode_var.get()

    if mode == "zoom":
        try:
            scale_factor = float(zoom_entry.get()) / 100
            if scale_factor <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Enter a valid zoom percentage!")
            return
    else:
        scale_factor = None

    output_folder = filedialog.askdirectory(title="Choose folder to save processed PDFs")
    if not output_folder:
        return

    def run():
        process_button.config(state=tk.DISABLED)
        try:
            for file in selected_files:
                zoom_pdf_content(file, output_folder, scale_factor, fabuchi=(mode == "fabuchi"))
            root.after(0, log_status, f"Processing complete! Files saved to: {output_folder}\n", "green")
        finally:
            root.after(0, lambda: process_button.config(state=tk.NORMAL))

    threading.Thread(target=run, daemon=True).start()


# =============================================================================
# SPLIT
# =============================================================================

def ask_folder_name(default_name):
    top = tk.Toplevel(root)
    top.title("Choose Output Folder Name")
    top.geometry("400x150")
    top.transient(root)
    top.grab_set()
    top.focus_force()

    top.update_idletasks()
    sw, sh = top.winfo_screenwidth(), top.winfo_screenheight()
    top.geometry(f"400x150+{(sw - 400) // 2}+{(sh - 150) // 2}")

    tk.Label(top, text="Enter name for output folder.\n"
                       "The folder will be saved alongside the PDF file.").pack(pady=10)
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


def split_pdf():
    if not check_libraries_loaded():
        return
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return
    if not check_all_pdfs():
        return
    if len(selected_files) > 1:
        log_status("Error: Select only one PDF for splitting!\n", "red")
        return

    pdf_path = selected_files[0]
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)
    pdf_document.close()

    input_ranges = simpledialog.askstring(
        "Page Ranges",
        f"Total pages: {total_pages}\n"
        f"Enter pairs of page numbers (space-separated).\n"
        f"Example: '1 10 11 20' → two parts: pages 1-10 and 11-20"
    )
    if not input_ranges:
        log_status("Error: No page ranges provided!\n", "red")
        return

    try:
        page_numbers = list(map(int, input_ranges.split()))
        if len(page_numbers) % 2 != 0:
            raise ValueError
    except ValueError:
        log_status("Error: Invalid input! Enter an even number of integers.\n", "red")
        return

    pdf_document = fitz.open(pdf_path)
    for i in range(0, len(page_numbers), 2):
        start, end = page_numbers[i], page_numbers[i + 1]
        if start < 1 or end > total_pages or start > end:
            log_status(f"Error: Invalid range {start}-{end}.\n", "red")
            pdf_document.close()
            return
    pdf_document.close()

    default_name = os.path.splitext(os.path.basename(pdf_path))[0] + "_split"
    folder_name = ask_folder_name(default_name)
    if not folder_name:
        return

    output_folder = os.path.join(os.path.dirname(pdf_path), folder_name)
    os.makedirs(output_folder, exist_ok=True)

    split_pdf_with_combined_output(pdf_path, page_numbers, output_folder)
    log_status(f"PDF split successfully! Files saved to: {output_folder}\n", "green")


def split_pdf_with_combined_output(pdf_path, page_ranges, output_folder):
    if not check_libraries_loaded():
        return

    pdf_document = fitz.open(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    individual_outputs = []
    merged_pdf = fitz.open()

    for i in range(0, len(page_ranges), 2):
        start, end = page_ranges[i], page_ranges[i + 1]
        new_pdf = fitz.open()
        new_pdf.insert_pdf(pdf_document, from_page=start - 1, to_page=end - 1)

        output_path = os.path.join(output_folder, f"{base_name}_{start}-{end}.pdf")
        new_pdf.save(output_path)
        new_pdf.close()
        individual_outputs.append(output_path)

        to_merge = fitz.open(output_path)
        merged_pdf.insert_pdf(to_merge)
        # Insert blank page if section has odd page count (for duplex printing)
        if (end - start + 1) % 2 != 0:
            blank = merged_pdf.new_page(width=to_merge[0].rect.width, height=to_merge[0].rect.height)
            blank.draw_rect(blank.rect)
        to_merge.close()

    merged_output_path = os.path.join(output_folder, f"{base_name}_merged.pdf")
    merged_pdf.save(merged_output_path)
    merged_pdf.close()
    pdf_document.close()

    return individual_outputs, merged_output_path


# =============================================================================
# MERGE
# =============================================================================

def merge_selected_pdfs():
    if not check_libraries_loaded():
        return
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return
    if not check_all_pdfs():
        return

    output_path = filedialog.asksaveasfilename(
        title="Save Merged PDF As",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")],
        initialfile="merged_output.pdf"
    )
    if not output_path:
        return

    try:
        merged = fitz.open()
        for file in selected_files:
            pdf = fitz.open(file)
            merged.insert_pdf(pdf)
            # Insert blank page if PDF has odd page count (for duplex printing)
            if len(pdf) % 2 != 0:
                rect = pdf[0].rect
                blank = merged.new_page(width=rect.width, height=rect.height)
                blank.draw_rect(blank.rect)
            pdf.close()

        merged.save(output_path)
        merged.close()
        log_status(f"Merged PDFs saved as: {os.path.basename(output_path)}\n", "green")

    except Exception as e:
        log_status(f"Error merging PDFs: {e}\n", "red")


# =============================================================================
# RESIZE TO A4
# =============================================================================

def resize_to_a4():
    if not check_libraries_loaded():
        return
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return
    if not check_all_pdfs():
        return

    output_folder = filedialog.askdirectory(title="Choose folder to save resized PDFs")
    if not output_folder:
        return

    A4_WIDTH, A4_HEIGHT = 595, 842

    for input_path in selected_files:
        try:
            pdf_document = fitz.open(input_path)
            new_pdf = fitz.open()

            for page_num in range(len(pdf_document)):
                new_page = new_pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
                new_page.show_pdf_page(new_page.rect, pdf_document, page_num, keep_proportion=True)

            name, ext = os.path.splitext(os.path.basename(input_path))
            output_path = os.path.join(output_folder, f"{name}_a4{ext}")
            new_pdf.save(output_path)
            pdf_document.close()
            new_pdf.close()
            log_status(f"Resized: {os.path.basename(output_path)}\n", "green")

        except Exception as e:
            log_status(f"Error resizing {os.path.basename(input_path)}: {e}\n", "red")


# =============================================================================
# WORD TO PDF
# =============================================================================

def convert_word_to_pdf():
    if not check_libraries_loaded():
        return
    if not selected_files:
        messagebox.showerror("Error", "No files selected!")
        return
    if not check_all_docx():
        return

    word_files = [f for f in selected_files if f.lower().endswith(".docx")]

    output_folder = filedialog.askdirectory(title="Choose folder to save converted PDFs")
    if not output_folder:
        return

    def run():
        for docx_path in word_files:
            success = False
            errors = []

            # Method 1: explicit output path
            try:
                output_path = os.path.join(
                    output_folder,
                    os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
                )
                convert_func(docx_path, output_path)
                root.after(0, log_status, f"Converted: {os.path.basename(docx_path)}\n", "green")
                success = True
            except Exception as e1:
                errors.append(f"Method 1: {e1}")

            # Method 2: folder only
            if not success:
                try:
                    convert_func(docx_path, output_folder)
                    root.after(0, log_status, f"Converted: {os.path.basename(docx_path)}\n", "green")
                    success = True
                except Exception as e2:
                    errors.append(f"Method 2: {e2}")

            # Method 3: with COM re-init (Windows)
            if not success:
                try:
                    import pythoncom
                    pythoncom.CoInitialize()
                    try:
                        convert_func(docx_path, output_folder)
                        root.after(0, log_status, f"Converted: {os.path.basename(docx_path)}\n", "green")
                        success = True
                    finally:
                        pythoncom.CoUninitialize()
                except Exception as e3:
                    errors.append(f"Method 3 (COM): {e3}")

            if not success:
                root.after(0, log_status, f"Failed to convert: {os.path.basename(docx_path)}\n", "red")
                root.after(0, log_status, f"  Last error: {errors[-1]}\n", "red")
                root.after(0, log_status, "  Try: close Word, run as admin, or check docx2pdf install\n", "blue")

    threading.Thread(target=run, daemon=True).start()


# =============================================================================
# GUI SETUP
# =============================================================================

try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# --- Color palette ---
BG          = "#f5f7fa"   # main background
SURFACE     = "#deeaf7"   # card/frame background (light blue)
SURFACE2    = "#ccdff0"   # slightly deeper blue (inputs, listbox)
BORDER      = "#a8c8e8"   # frame borders
ACCENT      = "#2a7abf"   # blue accent
ACCENT_HOV  = "#1e5f99"   # accent hover
ACCENT_DIM  = "#164d7a"   # accent pressed
FG          = "#111111"   # primary text
FG_DIM      = "#555555"   # secondary text
SUCCESS     = "#2e7d32"
ERROR       = "#c0392b"
INFO        = "#1565c0"

FONT_UI     = ("Segoe UI", 10)
FONT_LABEL  = ("Segoe UI", 9)
FONT_TITLE  = ("Segoe UI Semibold", 10)
FONT_MONO   = ("Consolas", 9)

root = tk.Tk()
root.title("The Ultimate PDF Tool 3000")
root.geometry("820x900")
root.resizable(True, True)
root.configure(bg=BG)
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# --- ttk Style ---
style = ttk.Style(root)
style.theme_use("clam")

style.configure(".",
    background=BG, foreground=FG,
    font=FONT_UI, borderwidth=0, focusthickness=0)

style.configure("TFrame", background=BG)
style.configure("Inner.TFrame", background=SURFACE)

style.configure("TLabel", background=BG, foreground=FG, font=FONT_LABEL)
style.configure("Dim.TLabel", background=SURFACE, foreground=FG_DIM, font=FONT_LABEL)
style.configure("Header.TLabel", background=BG, foreground=FG, font=FONT_TITLE)

# LabelFrame as a dark card with amber title
style.configure("Card.TLabelframe",
    background=SURFACE, relief="flat",
    borderwidth=1, bordercolor=BORDER,
    labelmargins=(8, 4))
style.configure("Card.TLabelframe.Label",
    background=SURFACE, foreground=ACCENT,
    font=FONT_TITLE, padding=(4, 0))
style.map("Card.TLabelframe", bordercolor=[("focus", ACCENT)])

# Primary accent button
style.configure("Accent.TButton",
    background=ACCENT, foreground="#ffffff",
    font=("Segoe UI Semibold", 10),
    relief="flat", borderwidth=0,
    padding=(10, 7))
style.map("Accent.TButton",
    background=[("active", ACCENT_HOV), ("pressed", ACCENT_DIM), ("disabled", BORDER)],
    foreground=[("disabled", FG_DIM)])

# Secondary ghost button
style.configure("Ghost.TButton",
    background=SURFACE2, foreground=FG,
    font=FONT_UI, relief="flat", borderwidth=0,
    padding=(10, 7))
style.map("Ghost.TButton",
    background=[("active", BORDER), ("pressed", "#222222")],
    foreground=[("active", FG)])

# Danger/clear button
style.configure("Danger.TButton",
    background=SURFACE, foreground=ERROR,
    font=FONT_UI, relief="flat", borderwidth=1,
    padding=(10, 7))
style.map("Danger.TButton",
    background=[("active", "#3a2424"), ("pressed", "#2a1a1a")],
    foreground=[("active", ERROR)])

style.configure("TRadiobutton",
    background=SURFACE, foreground=FG,
    font=FONT_LABEL, focusthickness=0,
    indicatorsize=16)
style.map("TRadiobutton",
    background=[("active", SURFACE)],
    foreground=[("active", ACCENT)])

style.configure("TEntry",
    fieldbackground=SURFACE2, foreground=FG,
    insertcolor=ACCENT, borderwidth=1,
    relief="flat", padding=(6, 4))
style.map("TEntry", bordercolor=[("focus", ACCENT)])

style.configure("TScrollbar",
    background=SURFACE2, troughcolor=SURFACE,
    borderwidth=0, arrowsize=12)
style.map("TScrollbar", background=[("active", BORDER)])

# --- Root layout ---
main_frame = ttk.Frame(root, padding=18)
main_frame.pack(fill="both", expand=True)
main_frame.columnconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)
main_frame.columnconfigure(2, weight=1)

# --- Header ---
header = tk.Frame(main_frame, bg=BG)
header.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 18))
header.columnconfigure(0, weight=1)

header_inner = tk.Frame(header, bg=BG)
header_inner.pack(anchor="center")

tk.Label(header_inner, text=" The Ultimate PDF Tool 3000 ", bg=BG, fg=ACCENT,
         font=("Segoe UI Black", 22, "bold")).pack(side="left")

# --- File Selection ---
file_frame = ttk.LabelFrame(main_frame, text="DOCUMENTS", style="Card.TLabelframe", padding=14)
file_frame.grid(row=1, column=0, columnspan=3, pady=(0, 14), sticky="ew")
file_frame.columnconfigure(0, weight=1)

btn_row = ttk.Frame(file_frame, style="Inner.TFrame")
btn_row.pack(fill="x", pady=(0, 10))

select_button = ttk.Button(btn_row, text="Browse", command=select_documents, style="Accent.TButton")
select_button.pack(side="left")

file_count_label = ttk.Label(btn_row, text="No files selected", style="Dim.TLabel")
file_count_label.pack(side="right", padx=(0, 4))

listbox_frame = ttk.Frame(file_frame, style="Inner.TFrame")
listbox_frame.pack(fill="both", expand=True)

file_listbox = tk.Listbox(
    listbox_frame, height=5, selectmode=tk.EXTENDED,
    bg=SURFACE2, fg=FG, selectbackground=ACCENT, selectforeground="#ffffff",
    font=FONT_MONO, relief="flat", borderwidth=0,
    highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
    activestyle="none")
scrollbar = ttk.Scrollbar(listbox_frame, orient="vertical", command=file_listbox.yview)
file_listbox.configure(yscrollcommand=scrollbar.set)
file_listbox.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# --- Operation panels row ---
# PDF Operations
tools_frame = ttk.LabelFrame(main_frame, text="PDF OPERATIONS", style="Card.TLabelframe", padding=14)
tools_frame.grid(row=2, column=0, padx=(0, 8), pady=(0, 14), sticky="nsew")

ttk.Button(tools_frame, text="Split PDF",   command=split_pdf,            style="Ghost.TButton").pack(fill="x", pady=(0, 6))
ttk.Button(tools_frame, text="Merge PDFs",  command=merge_selected_pdfs,  style="Ghost.TButton").pack(fill="x", pady=(0, 6))
ttk.Button(tools_frame, text="Resize to A4",command=resize_to_a4,         style="Ghost.TButton").pack(fill="x")

# Zoom & Process
zoom_frame = ttk.LabelFrame(main_frame, text="ZOOM & PROCESS", style="Card.TLabelframe", padding=14)
zoom_frame.grid(row=2, column=1, padx=(8, 8), pady=(0, 14), sticky="nsew")

mode_var = tk.StringVar(value="zoom")

ttk.Radiobutton(zoom_frame, text="Custom zoom", variable=mode_var, value="zoom").pack(anchor="w")

zoom_input_frame = ttk.Frame(zoom_frame, style="Inner.TFrame")
zoom_input_frame.pack(fill="x", padx=(18, 0), pady=(6, 0))
ttk.Label(zoom_input_frame, text="Amount:", style="Dim.TLabel").pack(side="left")
zoom_entry = ttk.Entry(zoom_input_frame, width=6)
zoom_entry.pack(side="left", padx=(6, 0))
zoom_entry.insert(0, "107")
ttk.Label(zoom_input_frame, text="%", style="Dim.TLabel").pack(side="left", padx=(3, 0))

ttk.Radiobutton(zoom_frame, text="Fabuchi preset", variable=mode_var, value="fabuchi").pack(anchor="w", pady=(10, 0))

process_button = ttk.Button(zoom_frame, text="Process PDFs", command=process_pdfs, style="Accent.TButton")
process_button.pack(fill="x", pady=(16, 0))

# Word Operations
word_frame = ttk.LabelFrame(main_frame, text="WORD OPERATIONS", style="Card.TLabelframe", padding=14)
word_frame.grid(row=2, column=2, padx=(8, 0), pady=(0, 14), sticky="nsew")

ttk.Button(word_frame, text="Word to PDF", command=convert_word_to_pdf, style="Ghost.TButton").pack(fill="x")

# --- Clear Selection ---
ttk.Button(main_frame, text="Clear Selection", command=clear_selection, style="Danger.TButton").grid(
    row=3, column=0, columnspan=3, pady=(0, 14), ipadx=20)

# --- Status ---
status_frame = ttk.LabelFrame(main_frame, text="LOG", style="Card.TLabelframe", padding=14)
status_frame.grid(row=4, column=0, columnspan=3, sticky="ew")

status_container = ttk.Frame(status_frame, style="Inner.TFrame")
status_container.pack(fill="both", expand=True)

status_text = tk.Text(
    status_container, height=8, state=tk.DISABLED, wrap=tk.WORD,
    font=FONT_MONO, bg=SURFACE2, fg=FG,
    relief="flat", borderwidth=0,
    highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER,
    insertbackground=ACCENT, padx=8, pady=6)
status_scrollbar = ttk.Scrollbar(status_container, orient="vertical", command=status_text.yview)
status_text.configure(yscrollcommand=status_scrollbar.set)
status_text.pack(side="left", fill="both", expand=True)
status_scrollbar.pack(side="right", fill="y")

status_text.tag_configure("green", foreground=SUCCESS)
status_text.tag_configure("red",   foreground=ERROR)
status_text.tag_configure("blue",  foreground=INFO)

ttk.Button(status_frame, text="Clear Log", style="Ghost.TButton",
           command=lambda: [status_text.config(state=tk.NORMAL),
                            status_text.delete(1.0, tk.END),
                            status_text.config(state=tk.DISABLED)]).pack(pady=(10, 0))

# =============================================================================
# STARTUP
# =============================================================================

root.after(100, log_status, "Loading libraries in background...\n", "blue")
threading.Thread(target=preload_libraries, daemon=True).start()
root.mainloop()