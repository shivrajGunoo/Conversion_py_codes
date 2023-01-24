from flask import Flask, request, send_file
from docx import Document
from io import BytesIO
from PyPDF2 import PdfFileWriter, PdfFileReader
import psutil

app = Flask(__name__)

def convert_word_to_pdf(file):
    """Converts a Microsoft Word document to a PDF.
    :param file: The Word document file.
    :return: A BytesIO object containing the PDF.
    """
    # Open the Word document
    doc = Document(file)

    # Create a new PDF file
    pdf_file = BytesIO()
    pdf_writer = PdfFileWriter()

    # Add each page of the Word document to the PDF
    for page in doc.paragraphs:
        pdf_writer.addPage(page.text)

    # Write the PDF to the BytesIO object
    pdf_writer.write(pdf_file)

    # Reset the BytesIO object to the beginning
    pdf_file.seek(0)
    return pdf_file

def check_memory_usage():
    process = psutil.Process(os.getpid())
    mem_info = process.memory_info()
    if mem_info.rss > threshold:
        return False
    else:
        return True

@app.route('/convert', methods=['POST'])
def convert():
    try:
        if not check_memory_usage():
            return "Error: Memory usage is too high."
        # Get the uploaded files
        files = request.files.getlist("file")
        pdf_files = []
        for file in files:
            pdf_file = convert_word_to_pdf(file)
            pdf_files.append(pdf_file)

        # Send the PDFs as a response
        return send_files(
            pdf_files,
            attachment_filename='converted.pdf',
            as_attachment=True,
            mimetype='application/pdf'
        )
    except Exception as e:
        return "An error occurred: {}".format(e)

if __name__ == '__main__':
    app.run(debug=True)
