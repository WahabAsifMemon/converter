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
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.table import WD_ALIGN_VERTICAL
from pptx.dml.color import RGBColor



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

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

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


@app.route('/pdf-to-ppt', methods=['POST'])
def upload_pdf_to_ppt():
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

            # Convert PDF to PowerPoint
            df_list = tabula.read_pdf(file_path, pages='all', multiple_tables=True)

            pptx_file_path = os.path.join(OUTPUT_FOLDER, f"{os.path.splitext(file.filename)[0]}.pptx")

            # Initialize presentation object
            prs = Presentation()

            # Loop through each DataFrame (assuming each DataFrame corresponds to a slide)
            for i, df in enumerate(df_list):
                slide_layout = prs.slide_layouts[i % len(prs.slide_layouts)]
                slide = prs.slides.add_slide(slide_layout)

                # Determine placeholder type to add a table
                placeholder = slide.placeholders[1]  # Adjust index based on your specific layout
                placeholder.text = f"Table {i + 1}"  # Set placeholder text

                # Calculate table dimensions and position
                left = Inches(1)
                top = Inches(1.5)
                width = Inches(8)
                height = Inches(5)

                # Add a table to the slide
                rows = len(df) + 1  # Including header row
                cols = len(df.columns)
                table = placeholder.insert_table(rows=rows, cols=cols, left=left, top=top, width=width, height=height)

                # Set table properties
                table.first_row = True  # Set first row as header
                table.horz_banding = True  # Enable horizontal banding (alternating row colors)

                # Populate table with data
                for row_idx, row in enumerate(df.itertuples(index=False)):
                    for col_idx, value in enumerate(row):
                        table.cell(row_idx + 1, col_idx).text = str(value)

                        # Adjust text alignment if needed
                        cell = table.cell(row_idx + 1, col_idx)
                        for paragraph in cell.text_frame.paragraphs:
                            paragraph.alignment = PP_ALIGN.CENTER  # Adjust alignment as needed
                            for run in paragraph.runs:
                                run.font.size = Pt(10)  # Adjust font size as needed
                                run.font.color.rgb = RGBColor(0, 0, 0)  # Adjust font color as needed

            # Save the PowerPoint presentation
            prs.save(pptx_file_path)

            return jsonify({'filename': f"{os.path.splitext(file.filename)[0]}.pptx"}), 200

        else:
            logging.error('Invalid file type, only PDF files are allowed')
            return jsonify({'error': 'Invalid file type, only PDF files are allowed'}), 400

    except Exception as e:
        logging.error(f'Error during file upload: {e}')
        return jsonify({'error': f'File upload failed: {str(e)}'}), 500

@app.route('/download_all')
def download_all():
    try:
        files = request.args.get('files').split(',')
        original_filename = request.args.get('filename', 'files')

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file in files:
                zip_file.write(os.path.join(OUTPUT_FOLDER, file), file)
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
