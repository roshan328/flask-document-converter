from flask import Flask, request, render_template, send_file
import os
from convert import docx_to_pdf, pdf_to_docx

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    conversion_type = request.form.get('conversion_type')
    if file and conversion_type:
        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(input_path)
        
        try:
            if conversion_type == 'docx_to_pdf':
                output_path = input_path.replace('.docx', '.pdf')
                docx_to_pdf(input_path, output_path)
            elif conversion_type == 'pdf_to_docx':
                output_path = input_path.replace('.pdf', '.docx')
                pdf_to_docx(input_path, output_path)
            else:
                return "Invalid conversion type", 400

            if os.path.isfile(output_path):
                return send_file(output_path, as_attachment=True)
            else:
                return "Conversion Failed: Output file not found", 500

        except Exception as e:
            return f"Conversion Failed: {e}", 500

    return "Conversion Failed: No file or conversion type provided", 400

if __name__ == '__main__':
    app.run(debug=True)







