from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
from pptx import Presentation
import pdfkit

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/convert', methods=['POST'])
def convert():
    try:
        # Check if the file is present
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        # Get the file from the request
        files = request.files.getlist('file')
        for file in files:
            # Check if the file is an Excel file
            if not allowed_file(file.filename):
                return jsonify({'error': 'Invalid file type. Only pptx files are supported.'}), 400
            
            # Save the file
            filename = secure_filename(file.filename)
            file.save(filename)
            
            # Open the PPT file
            prs = Presentation(filename)
            
            # Convert the PPT file to PDF
            pdf_file = filename.replace(".pptx", ".pdf")
            prs.save(pdf_file)
            
            return jsonify({'message': 'File converted successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)