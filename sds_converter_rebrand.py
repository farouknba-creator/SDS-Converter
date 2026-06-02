import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import json
import requests
from PyPDF2 import PdfReader
from docx import Document

# -------------------------------------------------------------------
# Your permanent company details – edit these once
# -------------------------------------------------------------------
COMPANY_NAME = "Your Company Name"
COMPANY_ADDRESS = "123 Chemical Lane, Industrial City"
COMPANY_PHONE = "+1 555 123 4567"
COMPANY_WEBSITE = "www.yourcompany.com"
# If you want to insert the logo programmatically, add its path here.
# For simplicity, embed the logo directly in template.docx header.
LOGO_PATH = None   # leave None if logo is already in the template

# -------------------------------------------------------------------
# DeepSeek settings
# -------------------------------------------------------------------
DEEPSEEK_API_KEY = ""   # paste your key, or leave blank to be prompted
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
MODEL = "deepseek-chat"

# Prompt to clean the SDS text – it removes the supplier block and returns only the body
CLEAN_PROMPT = """
You are given the full text of a Safety Data Sheet (SDS) that starts with a supplier identification block (company name, address, phone, logo text, etc.).
Remove all lines that belong to that supplier identification block. Keep everything else exactly as it is, including section numbers, headings, and all technical content.
Return ONLY the cleaned text, with no additional commentary.

Original SDS text:
---
{sds_text}
---
Cleaned SDS body:
"""

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
# Core functions
# -------------------------------------------------------------------
def extract_text_from_pdf(pdf_path):
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def call_deepseek(prompt, api_key):
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": "You are a precise text editor. Output only the requested text."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.0,
        "max_tokens": 4096
    }
    resp = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]

def clean_sds_text(sds_text, api_key):
    prompt = CLEAN_PROMPT.format(sds_text=sds_text)
    cleaned = call_deepseek(prompt, api_key)
    # Remove possible markdown fences
    cleaned = cleaned.strip().removeprefix("```").rstrip("```").strip()
    return cleaned

def fill_template(product_name, cleaned_body, output_path):
    template_path = resource_path("template.docx")
    if not os.path.exists(template_path):
        raise FileNotFoundError("template.docx not found in app folder")
    doc = Document(template_path)

    # Replace placeholders in paragraphs
    for para in doc.paragraphs:
        if "{{product_name}}" in para.text:
            para.text = para.text.replace("{{product_name}}", product_name)
        if "{{sds_content}}" in para.text:
            # If the placeholder is on its own line, we replace the whole paragraph with the body
            # But keep the paragraph style by inserting runs
            if para.text.strip() == "{{sds_content}}":
                para.clear()
                para.add_run(cleaned_body)
            else:
                para.text = para.text.replace("{{sds_content}}", cleaned_body)

    # Also check tables (just in case)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if "{{product_name}}" in para.text:
                        para.text = para.text.replace("{{product_name}}", product_name)
                    if "{{sds_content}}" in para.text:
                        if para.text.strip() == "{{sds_content}}":
                            para.clear()
                            para.add_run(cleaned_body)
                        else:
                            para.text = para.text.replace("{{sds_content}}", cleaned_body)

    doc.save(output_path)

def process_sds(pdf_path, output_path, api_key, product_name=None, progress_callback=None):
    # Use filename as product name if not provided
    if not product_name:
        product_name = os.path.splitext(os.path.basename(pdf_path))[0]

    if progress_callback: progress_callback("Extracting text from PDF...")
    raw_text = extract_text_from_pdf(pdf_path)

    if progress_callback: progress_callback("Removing supplier details (DeepSeek)...")
    cleaned = clean_sds_text(raw_text, api_key)

    if progress_callback: progress_callback("Generating Word document...")
    fill_template(product_name, cleaned, output_path)
    return product_name

# -------------------------------------------------------------------
# GUI
# -------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SDS Rebrander")
        self.geometry("550x350")
        self.resizable(False, False)
        self.api_key = DEEPSEEK_API_KEY

        if not self.api_key:
            self.show_api_key_prompt()
        else:
            self.build_main_ui()

    def show_api_key_prompt(self):
        for widget in self.winfo_children():
            widget.destroy()
        tk.Label(self, text="Enter your DeepSeek API key:", font=("Arial", 12)).pack(pady=20)
        self.api_entry = tk.Entry(self, width=50, show="*")
        self.api_entry.pack(pady=5)
        tk.Button(self, text="Save & Continue", command=self.save_api_key).pack(pady=10)

    def save_api_key(self):
        key = self.api_entry.get().strip()
        if not key:
            messagebox.showerror("Error", "API key cannot be empty.")
            return
        self.api_key = key
        for widget in self.winfo_children():
            widget.destroy()
        self.build_main_ui()

    def build_main_ui(self):
        tk.Label(self, text="SDS Rebrander", font=("Arial", 16, "bold")).pack(pady=10)
        tk.Label(self, text="Convert any Safety Data Sheet to your company format.\n"
                           "Supplier details are removed and replaced with your branding.",
                 wraplength=450, justify="left").pack(pady=5)

        # File selection
        frame = tk.Frame(self)
        frame.pack(pady=5)
        self.file_label = tk.Label(frame, text="No PDF selected", fg="gray", width=45, anchor="w")
        self.file_label.pack(side="left", padx=5)
        tk.Button(frame, text="Choose PDF...", command=self.select_pdf).pack(side="left")

        # Product name (optional manual override)
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

        self.pdf_path = None

    def select_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.file_label.config(text=os.path.basename(path))
            self.convert_btn.config(state="normal")
            # Auto‑fill product name from filename (without extension)
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
            process_sds(self.pdf_path, output_path, self.api_key, product_name,
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
