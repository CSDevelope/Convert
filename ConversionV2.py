from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
import tempfile
from docx import Document
from fpdf import FPDF
import pandas as pd
from PIL import Image
import win32com.client
from pywintypes import com_error
import pythoncom

app = Flask(__name__)

PDF_FOLDER = './pdfs'
os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs('./fonts', exist_ok=True)

def convert_docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'fonts/DejaVuSans.ttf')
    pdf.set_font('DejaVu', size=10)

    for para in doc.paragraphs:
        if para.text.strip():
            pdf.multi_cell(190, 10, para.text)
            pdf.ln(5)

    pdf.output(pdf_path)

def convert_image_to_pdf(image_path, pdf_path):
    image = Image.open(image_path)
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '', 'fonts/DejaVuSans.ttf')
    pdf.set_font('DejaVu', size=10)
    pdf.image(image_path, x=10, y=10, w=190)
    pdf.output(pdf_path)

def convert_excel_to_pdf(excel_path, pdf_path):
    pythoncom.CoInitialize()  # Initialize COM library
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(excel_path)
        temp_pdf_path = os.path.join(tempfile.gettempdir(), os.path.basename(pdf_path))
        wb.WorkSheets.Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, temp_pdf_path)
        # Move the temporary PDF to the desired location
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        os.rename(temp_pdf_path, pdf_path)
    except com_error as e:
        print('Excel to PDF conversion failed:', e)
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()  # Uninitialize COM library

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file part"
        file = request.files['file']
        if file.filename == '':
            return "No selected file"

        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1]) as tmp_file:
            filepath = tmp_file.name
            file.save(filepath)

        pdf_filename = secure_filename(file.filename.rsplit('.', 1)[0] + '.pdf')
        pdf_filepath = os.path.join(PDF_FOLDER, pdf_filename)
        file_extension = os.path.splitext(file.filename)[1].lower()

        try:
            if file_extension == '.docx':
                convert_docx_to_pdf(filepath, pdf_filepath)
            elif file_extension in ['.png', '.jpg', '.jpeg']:
                convert_image_to_pdf(filepath, pdf_filepath)
            elif file_extension in ['.xls', '.xlsx']:
                convert_excel_to_pdf(filepath, pdf_filepath)
            else:
                return "Unsupported file type"
        finally:
            if os.path.exists(filepath):
                os.remove(filepath)

        return send_file(pdf_filepath, as_attachment=True)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
