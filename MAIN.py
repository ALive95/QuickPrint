import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from ttkthemes import ThemedTk  # Import ThemedTk
import tksvg
import fitz  # PyMuPDF


def fabuchi_clip_rect(width, height):
    """Define the custom Fabuchi scaling rectangle."""
    return fitz.Rect(
        width * 0.05,  # Placeholder left bound
        height * 0.05,  # Placeholder top bound
        width * 0.95,  # Placeholder right bound
        height * 0.95  # Placeholder bottom bound
    )


def zoom_pdf_content(input_path, output_folder, scale_factor=None, fabuchi=False):
    """Zoom PDF content while maintaining original page size or apply Fabuchi scaling."""
    os.makedirs(output_folder, exist_ok=True)
    name, ext = os.path.splitext(os.path.basename(input_path))
    suffix = "_fabuchi" if fabuchi else "_zoomed"
    output_path = os.path.join(output_folder, f"{name}{suffix}{ext}")

    def update_status_text(text, color):
        """Update status text safely from a separate thread."""
        status_text.config(state=tk.NORMAL)  # Enable editing
        status_text.insert(tk.END, text, color)  # Insert text with color tag
        status_text.config(state=tk.DISABLED)  # Lock editing
        status_text.yview(tk.END)  # Auto-scroll to the latest entry

    try:
        pdf_document = fitz.open(input_path)
        new_pdf = fitz.open()

        for page_num in range(len(pdf_document)):
            original_page = pdf_document[page_num]
            new_page = new_pdf.new_page(width=original_page.rect.width,
                                        height=original_page.rect.height)

            width, height = original_page.rect.width, original_page.rect.height

            if fabuchi:
                clip_rect = fabuchi_clip_rect(width, height)
            else:
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

        # Update status text
        root.after(0, update_status_text, f"Processed: {os.path.basename(output_path)}\n", "green")  # Schedule update

    except Exception as e:
        # Update status text on error
        root.after(0, update_status_text, f"Error processing {input_path}: {e}\n", "red")  # Schedule update


def process_pdfs_in_thread():
    """Runs the PDF processing in a separate thread."""
    process_button.config(state=tk.DISABLED)  # Disable button to prevent multiple clicks
    try:
        process_pdfs()  # Run the processing function
    finally:
        process_button.config(state=tk.NORMAL)  # Re-enable button after processing is done


def process_pdfs():
    """Processes selected PDFs with chosen mode and zoom level."""
    if not selected_files:
        messagebox.showerror("Error", "No PDFs selected!")
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

    output_folder = os.path.join(os.getcwd(), "zoomies")

    # Process PDFs without updating status for every file
    for file in selected_files:
        zoom_pdf_content(file, output_folder, scale_factor, fabuchi=(mode == "fabuchi"))

    # Update status once after processing
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, "Processing complete!\n", "green")
    status_text.config(state=tk.DISABLED)


def select_pdfs():
    """Opens a file dialog to select PDFs."""
    global selected_files
    selected_files = filedialog.askopenfilenames(
        title="Select PDFs",
        filetypes=[("PDF Files", "*.pdf")]
    )
    file_listbox.delete(0, tk.END)
    for f in selected_files:
        file_listbox.insert(tk.END, os.path.basename(f))


from tkinter import simpledialog, messagebox
import fitz  # PyMuPDF
import os


def split_pdf():
    """Splits selected PDFs into parts based on user-defined page ranges, ensuring validity before processing."""

    if not selected_files:
        messagebox.showerror("Error", "No PDFs selected!")
        return

    if len(selected_files) > 1:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Error: Select only one PDF for splitting!\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)
        return
    pdf_path = selected_files[0]
    pdf_document = fitz.open(pdf_path)
    total_pages = len(pdf_document)

    # Ask for page ranges and show total pages
    input_ranges = simpledialog.askstring(
        "Page Ranges",
        f"Total pages: {total_pages}\nEnter 2n numbers (space-separated) defining the page ranges:"
    )

    if not input_ranges:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Error: No page ranges provided!\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)
        return

    # Convert input to a list of integers
    try:
        page_numbers = list(map(int, input_ranges.split()))
        if len(page_numbers) % 2 != 0:
            raise ValueError
    except ValueError:
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Error: Invalid input! Enter an even number of integers.\n", "red")
        status_text.config(state=tk.DISABLED)
        status_text.yview(tk.END)
        return

    # Validate all page ranges before processing
    for i in range(0, len(page_numbers), 2):
        start, end = page_numbers[i], page_numbers[i + 1]
        if start < 1 or end > total_pages or start > end:
            status_text.config(state=tk.NORMAL)
            status_text.insert(tk.END, f"Error: Invalid range {start}-{end}.\n", "red")
            status_text.config(state=tk.DISABLED)
            status_text.yview(tk.END)
            return  # Stop execution if any range is invalid

    base_name, ext = os.path.splitext(os.path.basename(pdf_path))
    output_folder = os.path.dirname(pdf_path)

    # Process page ranges
    for i in range(0, len(page_numbers), 2):
        start, end = page_numbers[i], page_numbers[i + 1]
        new_pdf = fitz.open()
        for page_num in range(start - 1, end):  # Convert 1-based to 0-based indexing
            new_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

        output_path = os.path.join(output_folder, f"{base_name}_{start}-{end}.pdf")
        new_pdf.save(output_path)
        new_pdf.close()

    pdf_document.close()

    # Success message
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, "PDF split successfully!\n", "green")
    status_text.config(state=tk.DISABLED)
    status_text.yview(tk.END)


# Use ThemedTk instead of Tk
root = ThemedTk(theme="adapta")  # Apply Adapta theme
root.title("PDF Rescaler")
root.geometry("450x550")
root.resizable(False, False)

# Main Frame
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill="both", expand=True)

# Select PDFs Section
file_frame = ttk.LabelFrame(main_frame, text="Select PDFs", padding=10)
file_frame.grid(row=0, column=0, columnspan=2, pady=3, sticky="ew")

select_button = ttk.Button(file_frame, text="📂 Select PDFs", command=select_pdfs)
select_button.pack(pady=5)
# Apply the theme to the listbox explicitly
style = ttk.Style()

file_listbox = tk.Listbox(file_frame, height=4, width=55)
file_listbox.pack(pady=3)

# Mode Selection - Using Grid for Side by Side Frames
zoom_frame = ttk.LabelFrame(main_frame, text="Zoom & Process", padding=10)
zoom_frame.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

split_frame = ttk.LabelFrame(main_frame, text="PDF Splitting", padding=10)
split_frame.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

# Zoom Mode Selection
mode_var = tk.StringVar(value="zoom")
ttk.Radiobutton(zoom_frame, text="Zoom (Custom %)", variable=mode_var, value="zoom").pack(anchor="w")
zoom_entry = ttk.Entry(zoom_frame, width=10)
zoom_entry.pack(pady=5)
zoom_entry.insert(0, "107")

ttk.Radiobutton(zoom_frame, text="Fabuchi (Preset Bounds)", variable=mode_var, value="fabuchi").pack(anchor="w")

# Process PDFs Button
process_button = ttk.Button(zoom_frame, text="🔍 Process PDFs", command=process_pdfs_in_thread)
process_button.pack(pady=10)

# Split PDFs Button
split_button = ttk.Button(split_frame, text="✂ Split PDFs", command=split_pdf)
split_button.pack(pady=10)

# Status Text
status_frame = ttk.LabelFrame(main_frame, text="Status", padding=10)
status_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

status_text = tk.Text(status_frame, height=6, width=55, state=tk.DISABLED, wrap=tk.WORD, bg="#f0f0f0",
                      font=("Arial", 10))
status_text.pack()

# Define color tags for green and red text
status_text.tag_configure("green", foreground="green")
status_text.tag_configure("red", foreground="red")

selected_files = []
root.mainloop()
