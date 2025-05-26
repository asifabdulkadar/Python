import os
from tkinter import Label
from tkinterdnd2 import DND_FILES, TkinterDnD
from docx2pdf import convert
from fpdf import FPDF
import pdfkit

# === File Conversion Functions ===
def convert_docx_to_pdf(filepath):
    try:
        convert(filepath, os.path.dirname(filepath))
        return "✅ DOCX converted to PDF!"
    except Exception as e:
        return f"❌ DOCX Error: {e}"

def convert_txt_to_pdf(filepath):
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.readlines()

        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)

        for line in content:
            pdf.multi_cell(0, 10, line.strip())

        output_path = filepath.replace(".txt", ".pdf")
        pdf.output(output_path)
        return "✅ TXT converted to PDF!"
    except Exception as e:
        return f"❌ TXT Error: {e}"

def convert_html_to_pdf(filepath):
    try:
        output_path = filepath.replace(".html", ".pdf").replace(".htm", ".pdf")
        pdfkit.from_file(filepath, output_path)
        return "✅ HTML converted to PDF!"
    except Exception as e:
        return f"❌ HTML Error: {e}"

# === Drag-and-Drop Handler ===
def handle_file_drop(event):
    raw_path = event.data.strip()
    # Handle curly braces for file paths with spaces
    if raw_path.startswith('{') and raw_path.endswith('}'):
        raw_path = raw_path[1:-1]
    file_path = raw_path.replace("\\", "/")

    # Lowercase for consistent extension matching
    ext = os.path.splitext(file_path)[-1].lower()

    if ext == ".docx":
        result = convert_docx_to_pdf(file_path)
    elif ext == ".txt":
        result = convert_txt_to_pdf(file_path)
    elif ext in [".html", ".htm"]:
        result = convert_html_to_pdf(file_path)
    else:
        result = "❌ Unsupported file format: Only .docx, .txt, .html"

    label.config(text=result)

# === UI Setup ===
root = TkinterDnD.Tk()
root.title("📄 Drag & Drop to PDF Converter")
root.geometry("480x240")
root.configure(bg="#f0f4f8")

label = Label(root, text="🖱️ Drop a .docx, .txt, or .html file here\n\nIt’ll magically become a PDF ✨",
              bg="#f0f4f8", fg="#222", font=("Segoe UI", 12), relief="solid", bd=1, width=50, height=6)
label.pack(pady=30)

label.drop_target_register(DND_FILES)
label.dnd_bind('<<Drop>>', handle_file_drop)

root.mainloop()
