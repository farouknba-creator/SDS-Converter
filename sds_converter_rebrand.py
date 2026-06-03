import tkinter as tk
from tkinter import filedialog, messagebox
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
# Rule-based supplier removal (FREE, OFFLINE)
# -------------------------------------------------------------------
def strip_supplier_section(text):
    """
    Remove supplier identification block by cutting from the start of the document
    to the end of SECTION 1 (and any content before SECTION 2).
    Keeps the rest of the SDS intact.
    """
    # Normalise line endings and case for regex
    cleaned = text.strip()
    
    # Find the start of SECTION 2 (everything after is safety data we want to keep)
    # We'll keep everything from SECTION 2 onward, plus any product name that may appear
    # before SECTION 1. But most of the time we just want from SECTION 2 to end.
    
    # Pattern: "SECTION 2:" or "Section 2:" or "2." style headings
    match_section2 = re.search(
        r'(SECTION\s*2\s*:|Section\s*2\s*:|2\.\s*HAZARDS\s*IDENTIFICATION)',
        cleaned,
        re.IGNORECASE
    )
    
    if match_section2:
        # Keep text from the start of SECTION 2 to the end
        body = cleaned[match_section2.start():]
    else:
        # Fallback: if we can't find Section 2, try to remove the first block
        # that looks like an address (multiple lines with postal codes, phone, etc.)
        lines = cleaned.split('\n')
        start_keeping = 0
        for i, line in enumerate(lines):
            if re.search(r'(SECTION\s*2|HAZARDS\s*IDENTIFICATION|COMPOSITION)', line, re.I):
                start_keeping = i
                break
        body = '\n'.join(lines[start_keeping:]) if start_keeping else cleaned

    # Remove any remaining supplier name/logo lines that might be at the top
    # (e.g., a company name line before Section 1)
    body = re.sub(r'^.*(Sigma[-\s]?Aldrich|Fisher|Merck|BASF|Dow|DuPont).*\n?', '', body, flags=re.I)
    
    return body.strip()

def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

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

def process_sds(pdf_path, output_path, product_name=None, progress_callback=None):
    if not product_name:
        product_name = os.path.splitext(os.path.basename(pdf_path))[0]

    if progress_callback: progress_callback("Extracting text from PDF...")
    raw_text = extract_text_from_pdf(pdf_path)

    if progress_callback: progress_callback("Removing supplier details (rule-based)...")
    cleaned = strip_supplier_section(raw_text)

    if progress_callback: progress_callback("Generating Word document...")
    fill_template(product_name, cleaned, output_path)
    return product_name

# -------------------------------------------------------------------
# GUI (unchanged from before, minus API key prompts)
# -------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SDS Rebrander (Free)")
        self.geometry("550x350")
        self.resizable(False, False)
        self.build_main_ui()

    def build_main_ui(self):
        tk.Label(self, text="SDS Rebrander (Free)", font=("Arial", 16, "bold")).pack(pady=10)
        tk.Label(self, text="Convert any Safety Data Sheet to your company format.\n"
                           "Supplier details are automatically removed.",
                 wraplength=450, justify="left").pack(pady=5)

        frame = tk.Frame(self)
        frame.pack(pady=5)
        self.file_label = tk.Label(frame, text="No PDF selected", fg="gray", width=45, anchor="w")
        self.file_label.pack(side="left", padx=5)
        tk.Button(frame, text="Choose PDF...", command=self.select_pdf).pack(side="left")

        name_frame = tk.Frame(self)
        name_frame.pack(pady=5)
        tk.Label(name_frame, text="Product name (blank = use filename):").pack(side="left")
        self.product_name_var = tk.StringVar()
        tk.Entry(name_frame, textvariable=self.product_name_var, width=30).pack(side="left", padx=5)

        self.convert_btn = tk.Button(self, text="Convert", state="disabled", command=self.convert)
        self.convert_btn.pack(pady=10)

        self.status = tk.Label(self, text="", fg="blue")
        self.status.pack(pady=5)

        self.pdf_path = None

    def select_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.file_label.config(text=os.path.basename(path))
            self.convert_btn.config(state="normal")
            base = os.path.splitext(os.path.basename(path))[0]
            self.product_name_var.set(base)

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
