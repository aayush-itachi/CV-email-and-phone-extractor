from flask import Flask, render_template, request, send_file
import re
import pandas as pd
from docx import Document
import PyPDF2
from io import BytesIO

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'pdf', 'docx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_info_from_pdf(pdf_file):
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ''
    for page in pdf_reader.pages:
        text += page.extract_text()
    email = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    phone = re.findall(r'\b\d{10}\b', text)
    return {'email': email, 'phone': phone, 'text': text}

def extract_info_from_docx(docx_file):
    doc = Document(BytesIO(docx_file))
    text = ''
    for paragraph in doc.paragraphs:
        text += paragraph.text
    email = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    phone = re.findall(r'\b\d{10}\b', text)
    return {'email': email, 'phone': phone, 'text': text}

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            if file.filename.endswith('.pdf'):
                info = extract_info_from_pdf(file)
            elif file.filename.endswith('.docx'):
                info = extract_info_from_docx(file.read())
            df = pd.DataFrame([info])
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False)
            writer.close()
            output.seek(0)
            return send_file(output, mimetype='application/vnd.ms-excel', as_attachment=True, download_name='cv_info.xls')
        else:
            return render_template('upload.html', message='Invalid file format.')
    return render_template('upload.html', message='')

if __name__ == '__main__':
    app.run(debug=True)
