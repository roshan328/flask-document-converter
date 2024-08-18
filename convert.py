import os
import pythoncom
import comtypes.client
from docx2pdf import convert as docx2pdf_convert
from comtypes import COMError

def docx_to_pdf(docx_path, pdf_path):
    """Convert DOCX to PDF using COM automation."""
    try:
        print(f"Input file path: {docx_path}")
        print(f"Output file path: {pdf_path}")

        if not os.path.isfile(docx_path):
            raise FileNotFoundError(f"The input file {docx_path} does not exist.")

        pythoncom.CoInitialize()  # Initialize COM library
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        pythoncom.CoUninitialize()  # Uninitialize COM library

        if not os.path.isfile(pdf_path):
            raise FileNotFoundError(f"The output file {pdf_path} was not created.")
    except COMError as e:
        print(f"COMError: {e}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        pythoncom.CoUninitialize()  # Ensure COM library is uninitialized

def pdf_to_docx(pdf_path, docx_path):
    """Convert PDF to DOCX using docx2pdf."""
    try:
        print(f"Converting PDF path: {pdf_path} to DOCX path: {docx_path}")
        docx2pdf_convert(pdf_path, docx_path)
        
        if not os.path.isfile(docx_path):
            raise FileNotFoundError(f"The output file {docx_path} was not created.")
    except Exception as e:
        print(f"Error: {e}")











