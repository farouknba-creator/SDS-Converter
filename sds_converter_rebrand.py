import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import sys
import re
from PyPDF2 import PdfReader
from docx import Document

# -------------------------------------------------------------------
# Your permanent company details
# -------------------------------------------------------------------
COMPANY_NAME = "Your Company Name"
COMPANY_ADDRESS = "123 Chemical Lane, Industrial City"
COMPANY_PHONE = "+1 555 123 4567"
COMPANY_WEBSITE = "www.yourcompany.com"
LOGO_PATH = None   # embed logo in template.docx

# -------------------------------------------------------------------
# Helper for PyInstaller bundled resources
# -------------------------------------------------------------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -------------------------------------------------------------------
# Safer supplier block removal – NEVER deletes everything
# -------------------------------------------------------------------
def strip_supplier_section(text):
    """
    Remove supplier identification block from the top of the SDS.
    Strategy:
    - Find the first significant safety heading (like "HAZARDS IDENTIFICATION"
      or "COMPOSITION") and keep everything from that point onward.
    - Also try to delete common supplier name lines if they appear before that.
    - If no heading is found, return the original text unchanged (better safe).
    """
    cleaned = text.strip()
    if not cleaned:
        return cleaned

    # List of section headings that indicate the start of the main SDS body
    # (case‑insensitive, whole words)
    safety_headings = [
        r'\bHAZARDS?\s*IDENTIFICATION\b',
        r'\bCOMPOSITION\s*\/?\s*INFORMATION\s+ON\s+INGREDIENTS\b',
        r'\bFIRST\s*AID\s+MEASURES\b',
        r'\bTOXICOLOGICAL\s+INFORMATION\b',
        r'\bSECTION\s*2\b',                     # GHS style "SECTION 2:"
        r'\b2\.?\s+HAZARDS\s*IDENTIFICATION\b'  # "2. Hazards Identification"
    ]

    # Find the earliest occurrence of any safety heading
    earliest = len(cleaned)
    for pattern in safety_headings:
        match = re.search(pattern, cleaned, re.IGNORECASE)
        if match and match.start() < earliest:
            earliest = match.start()

    if earliest < len(cleaned):
        # Keep everything from the safety heading onward
        body = cleaned[earliest:].strip()
    else:
        # No safety heading found – keep all text (don't destroy it)
        body = cleaned

    # Remove any remaining supplier name lines at the very top (just in case)
    # These are typical well‑known chemical suppliers. Extend with your own.
    supplier_patterns = [
        r'^.*(Sigma[-\s]?Aldrich|Fisher\s*Scientific|Merck\s*KGaA|BASF\s*SE|Dow\s*Chemical|DuPont|Thermo\s*Fisher).*\n?',
        r'^.*(www\.|http).*\n?',            # web addresses
        r'^.*(Tel|Phone|Fax|Emergency).*\n?'
    ]
    for pat in supplier_patterns:
        body = re.sub(pat, '', body, flags=re.IGNORECASE | re.MULTILINE)

    return body.strip()

# -------------------------------------------------------------------
# PDF text extraction
# -------------------------------------------------------------------
def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text.strip()

# -------------------------------------------------------------------
# Template filling
# -------------------------------------------------------------------
def fill_template(product_name, sds_body, output_path):
    template_path = resource_path("template.docx")
    if not os.path.exists(template_path):
        raise FileNotFoundError("template.docx not found in app folder")
    doc = Document(template_path)

    for para in doc.paragraphs:
        if "{{product_name}}" in para.text:
            para.text = para.text.replace("{{product_name}}", product_name)
        if "{{sds_content}}" in para.text:
            if para.text.strip() == "{{sds_content}}":
                para.clear()
                para.add_run(sds_body)
            else:
                para.text = para.text.replace("{{sds_content}}", sds_body)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if "{{product_name}}" in para.text:
                        para.text = para.text.replace("{{product_name}}", product_name)
                    if "{{sds_content}}" in para.text:
                        if para.text.strip() == "{{sds_content}}":
                            para.clear()
                            para.add_run(sds_body)
                        else:
                            para.text = para.text.replace("{{sds_content}}", sds_body)

    doc.save(output_path)

# -------------------------------------------------------------------
# Main processing
# -------------------------------------------------------------------
def process_sds(pdf_path, output_path, product_name=None, progress_callback=None):
    if not product_name:
        product_name = os.path.splitext(os.path.basename(pdf_path))[0]

    if progress_callback: progress_callback("Extracting text from PDF...")
    raw_text = extract_text_from_pdf(pdf_path)

    # Save raw text for debugging (sidecar file)
    txt_path = output_path.replace('.docx', '_raw_text.txt')
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(raw_text)

    if not raw_text:
        raise ValueError(
            "No text could be extracted from this PDF.\n\n"
            "The PDF is likely a scanned image (no digital text layer).\n"
            "A raw text file (empty) was saved next to the output for inspection."
        )

    if progress_callback: progress_callback("Removing supplier details...")
    cleaned = strip_supplier_section(raw_text)

    if not cleaned:
        # This should never happen with the new strip function, but just in case
        cleaned = raw_text  # fallback to original

    if progress_callback: progress_callback("Generating Word document...")
    fill_template(product_name, cleaned, output_path)
    return product_name

# -------------------------------------------------------------------
# GUI with diagnostic text box
# -------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SDS Rebrander (Free)")
        self.geometry("650x550")
        self.resizable(True, True)
        self.build_main_ui()

    def build_main_ui(self):
        tk.Label(self, text="SDS Rebrander", font=("Arial", 16, "bold")).pack(pady=10)
        tk.Label(self, text="Supplier details are automatically removed.\n"
                           "Extracted text preview below (for diagnostics).",
                 wraplength=500, justify="left").pack(pady=5)

        # File selection
        frame = tk.Frame(self)
        frame.pack(pady=5)
        self.file_label = tk.Label(frame, text="No PDF selected", fg="gray", width=45, anchor="w")
        self.file_label.pack(side="left", padx=5)
        tk.Button(frame, text="Choose PDF...", command=self.select_pdf).pack(side="left")

        # Product name
        name_frame = tk.Frame(self)
        name_frame.pack(pady=5)
        tk.Label(name_frame, text="Product name (blank = use filename):").pack(side="left")
        self.product_name_var = tk.StringVar()
        tk.Entry(name_frame, textvariable=self.product_name_var, width=30).pack(side="left", padx=5)

        # Convert button
        self.convert_btn = tk.Button(self, text="Convert", state="disabled", command=self.convert)
        self.convert_btn.pack(pady=10)

        self.status = tk.Label(self, text="", fg="blue")
        self.status.pack(pady=5)

        # Diagnostic text area
        tk.Label(self, text="Extracted text preview (first 2000 chars):", anchor="w").pack(pady=(10,0))
        self.preview = scrolledtext.ScrolledText(self, height=10, width=70, state="disabled")
        self.preview.pack(padx=10, pady=5)

        self.pdf_path = None

    def select_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.file_label.config(text=os.path.basename(path))
            self.convert_btn.config(state="normal")
            base = os.path.splitext(os.path.basename(path))[0]
            self.product_name_var.set(base)
            # Show a quick preview of raw extracted text
            try:
                raw = extract_text_from_pdf(path)
                self.preview.config(state="normal")
                self.preview.delete(1.0, tk.END)
                if raw:
                    self.preview.insert(tk.END, raw[:2000] + ("..." if len(raw) > 2000 else ""))
                else:
                    self.preview.insert(tk.END, "⚠️ WARNING: No extractable text! This PDF is probably scanned.")
                self.preview.config(state="disabled")
            except Exception as e:
                self.preview.config(state="normal")
                self.preview.delete(1.0, tk.END)
                self.preview.insert(tk.END, f"Error reading PDF: {e}")
                self.preview.config(state="disabled")

    def convert(self):
        if not self.pdf_path:
            return
        output_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                   filetypes=[("Word Document", "*.docx")],
                                                   initialfile=self.product_name_var.get() + ".docx")
        if not output_path:
            return
        self.convert_btn.config(state="disabled")
        self.status.config(text="Working...")
        self.update()

        try:
            product_name = self.product_name_var.get().strip() or None
            process_sds(self.pdf_path, output_path, product_name,
                        progress_callback=lambda msg: self.status.config(text=msg))
            self.status.config(text=f"Done! Saved to {output_path}")
            messagebox.showinfo("Success", f"Converted file saved to:\n{output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed:\n{str(e)}")
            self.status.config(text="Error occurred.")
        finally:
            self.convert_btn.config(state="normal")

if __name__ == "__main__":
    app = App()
    app.mainloop()
