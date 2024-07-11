import os
import logging
from flask import Flask, request, send_file, render_template, jsonify
from pdf2image import convert_from_path
import zipfile
import io
import datetime
from pdf2docx import Converter
import tabula
import pandas as pd
import fitz  # PyMuPDF library for PDF to PDF/A conversion
from docx2pdf import convert as docx2pdf_convert
from werkzeug.utils import secure_filename
import subprocess
from pptx import Presentation
from pptx.util import Inches

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s %(message)s')


def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                os.rmdir(file_path)
        except Exception as e:
            logging.error(f'Failed to delete {file_path}. Reason: {e}')


@app.route('/')
def home():
    return render_template('index.html')

@app.route('/pdf-to-jpg')
def pdf_to_jpg():
    return render_template('pdf-to-jpg.html')

@app.route('/pdf-to-word')
def pdf_to_word():
    return render_template('pdf-to-word.html')

@app.route('/pdf-to-excel')
def pdf_to_excel():
    return render_template('pdf-to-excel.html')

@app.route('/pdf-to-ppt')
def pdf_to_ppt():
    return render_template('pdf-to-ppt.html')

@app.route('/pdf-to-pdfa')
def pdf_to_pdfa():
    return render_template('pdf-to-pdfa.html')



ALLOWED_EXTENSIONS = {'pdf'}
ALLOWED_EXTENSIONS_DOCX = {'docx'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


# START PDF TO JPG
@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            logging.error('No file part in the request')
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']

        if file.filename == '':
            logging.error('No selected file')
            return jsonify({'error': 'No selected file'}), 400

        if file and allowed_file(file.filename):
            clear_folder(UPLOAD_FOLDER)
            clear_folder(OUTPUT_FOLDER)

            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            images = convert_from_path(file_path)
            image_files = []

            for i, image in enumerate(images):
                image_filename = f'page_{i + 1}.jpg'
                image_path = os.path.join(OUTPUT_FOLDER, image_filename)
                image.save(image_path, 'JPEG')
                image_files.append(image_filename)

            original_filename = os.path.splitext(file.filename)[0]

            return jsonify({'images': image_files, 'original_filename': original_filename})

        else:
            logging.error('Invalid file type, only PDF files are allowed')
            return jsonify({'error': 'Invalid file type, only PDF files are allowed'}), 400

    except Exception as e:
        logging.error(f'Error during file upload: {e}')
        return jsonify({'error': f'File upload failed: {str(e)}'}), 500
# END PDF TO JPG

# START PDF TO WORD
@app.route('/upload-pdf-to-word', methods=['POST'])
def upload_pdf_to_word():
    try:
        if 'file' not in request.files:
            logging.error('No file part in the request')
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']

        if file.filename == '':
            logging.error('No selected file')
            return jsonify({'error': 'No selected file'}), 400

        if file and allowed_file(file.filename):
            clear_folder(UPLOAD_FOLDER)
            clear_folder(OUTPUT_FOLDER)

            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)

            # Convert PDF to Word
            docx_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(file.filename)[0]}.docx")
            cv = Converter(file_path)
            cv.convert(docx_file_path)
            cv.close()

            return jsonify({'filename': f"{os.path.splitext(file.filename)[0]}.docx"}), 200

        else:
            logging.error('Invalid file type, only PDF files are allowed')
            return jsonify({'error': 'Invalid file type, only PDF files are allowed'}), 400

    except Exception as e:
        logging.error(f'Error during file upload: {e}')
        return jsonify({'error': f'File upload failed: {str(e)}'}), 500
# END PDF TO WORD


# START PDF TO EXCEL
@app.route('/upload-pdf-to-excel', methods=['POST'])
def upload_pdf_to_excel():
    try:
        if 'file' not in request.files:
            logging.error('No file part in the request')
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']

        if file.filename == '':
            logging.error('No selected file')
            return jsonify({'error': 'No selected file'}), 400

        if file and allowed_file(file.filename):
            clear_folder(UPLOAD_FOLDER)
            clear_folder(OUTPUT_FOLDER)

            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)

            # Convert PDF to Excel
            df_list = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
            excel_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(file.filename)[0]}.xlsx")

            with pd.ExcelWriter(excel_file_path) as writer:
                for i, df in enumerate(df_list):
                    df.to_excel(writer, sheet_name=f'Sheet{i+1}', index=False)

            return jsonify({'filename': f"{os.path.splitext(file.filename)[0]}.xlsx"}), 200

        else:
            logging.error('Invalid file type, only PDF files are allowed')
            return jsonify({'error': 'Invalid file type, only PDF files are allowed'}), 400

    except Exception as e:
        logging.error(f'Error during file upload: {e}')
        return jsonify({'error': f'File upload failed: {str(e)}'}), 500
# END PDF TO EXCEL


# START PDF TO PDF/A
@app.route('/upload-pdf-to-pdfa', methods=['POST'])
def upload_pdf_to_pdfa():
    try:
        if 'file' not in request.files:
            logging.error('No file part in the request')
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']

        if file.filename == '':
            logging.error('No selected file')
            return jsonify({'error': 'No selected file'}), 400

        if file and allowed_file(file.filename):
            clear_folder(UPLOAD_FOLDER)
            clear_folder(OUTPUT_FOLDER)

            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)

            # Convert PDF to PDF/A
            pdfa_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(file.filename)[0]}_pdfa.pdf")
            convert_to_pdfa(file_path, pdfa_file_path)

            return jsonify({'filename': f"{os.path.splitext(file.filename)[0]}_pdfa.pdf"}), 200

        else:
            logging.error('Invalid file type, only PDF files are allowed')
            return jsonify({'error': 'Invalid file type, only PDF files are allowed'}), 400

    except Exception as e:
        logging.error(f'Error during file upload: {e}')
        return jsonify({'error': f'File upload failed: {str(e)}'}), 500
# END PDF TO PDF/A


@app.route('/upload-pdf-to-ppt', methods=['POST'])
def upload_pdf_to_ppt():
    try:
        if 'file' not in request.files:
            logging.error('No file part in the request')
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']

        if file.filename == '':
            logging.error('No selected file')
            return jsonify({'error': 'No selected file'}), 400

        if file and allowed_file(file.filename, ALLOWED_EXTENSIONS):
            clear_folder(UPLOAD_FOLDER)
            clear_folder(OUTPUT_FOLDER)

            file_path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
            file.save(file_path)

            # Convert PDF to images (one image per page)
            images = convert_from_path(file_path, 300)  # 300 DPI for good quality

            # Create a PowerPoint presentation
            ppt_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(file.filename)[0]}.pptx")
            prs = Presentation()

            for i, image in enumerate(images):
                slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank slide layout
                slide.shapes.add_picture(io.BytesIO(image._repr_png_()), Inches(0), Inches(0), width=Inches(10))

            prs.save(ppt_file_path)

            return jsonify({'filename': f"{os.path.splitext(file.filename)[0]}.pptx"}), 200

        else:
            logging.error('Invalid file type, only PDF or PPTX files are allowed')
            return jsonify({'error': 'Invalid file type, only PDF or PPTX files are allowed'}), 400

    except Exception as e:
        logging.error(f'Error during file upload: {e}')
        return jsonify({'error': f'File upload failed: {str(e)}'}), 500



# CONVERTING PDF TO PDF/A
def convert_to_pdfa(input_path, output_path):
    # Using PyMuPDF to convert PDF to PDF/A
    try:
        doc = fitz.open(input_path)
        doc_pdf = fitz.open()
        doc_pdf.insert_pdf(doc)
        doc_pdf.save(output_path)
        doc_pdf.close()
    except Exception as e:
        logging.error(f'Error converting to PDF/A: {e}')
        raise
# CONVERTING PDF TO PDF/A



@app.route('/download_all')
def download_all():
    try:
        image_files = request.args.get('images').split(',')
        original_filename = request.args.get('filename', 'images')

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for image in image_files:
                zip_file.write(os.path.join(OUTPUT_FOLDER, image), image)
        zip_buffer.seek(0)

        # Constructing the zip filename based on the original filename
        timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        zip_filename = f'{original_filename}_{timestamp}.zip'

        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name=zip_filename)
    except Exception as e:
        logging.error(f'Error during file download: {e}')
        return jsonify({'error': f'File download failed: {str(e)}'}), 500

@app.route('/download')
def download():
    filename = request.args.get('filename')
    if filename:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({'error': 'Filename not provided'}), 400

if __name__ == '__main__':
    app.run(debug=True)
