from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
import openpyxl
import pdfkit

app = Flask(__name__)

ALLOWED_EXTENSIONS = {'xlsx'}

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
                return jsonify({'error': 'Invalid file type. Only xlsx files are supported.'}), 400
            
            # Save the file
            filename = secure_filename(file.filename)
            file.save(filename)
            
            # Open the Excel file
            workbook = openpyxl.load_workbook(filename)
            
            # Convert the Excel file to PDF
            pdfkit.from_file(filename, '{}.pdf'.format(filename))
            
            return jsonify({'message': 'File converted successfully'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
In this version, I've added a function allowed_file(filename) which checks if the file has a valid extension. The function checks if the filename has a "." in it, and then checks if the file extension (the part after the last ".") is in the list of allowed extensions. The list of allowed extensions is defined as a global variable ALLOWED_EXTENSIONS which is a set of strings containing the allowed file extensions in this case only 'xlsx'.

In the convert method, I've added a check for the file format using this function if not allowed_file(file.filename): if the file format is not valid the code will return an error message and a status code 400.

This way, the code will only proceed with the conversion if the file is an Excel file with the correct format, and will return an error message if the file is of the wrong format.




