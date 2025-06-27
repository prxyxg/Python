from pdf2docx import Converter
import os
 
# ✅ Set your input PDF and output DOCX paths here
input_pdf_path = 'D:\\TEST\\FDD Cir 01-2024 - Grant for Equity Market Singapore (GEMS) Scheme - 2 Jan 2024 3.pdf'
output_docx_path = 'D:\\TEST\\Test.docx'
 
def convert_pdf_to_docx(pdf_path, docx_path):
    # Check if PDF file exists
    if not os.path.isfile(pdf_path):
        print(f"Error: PDF file '{pdf_path}' does not exist.")
        return
 
    try:
        print(f"Converting '{pdf_path}' to '{docx_path}'...")
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        print("Conversion completed successfully.")
    except Exception as e:
        print(f"Conversion failed: {e}")
 
# Run the conversion
convert_pdf_to_docx(input_pdf_path, output_docx_path)