import cv2
import os
import comtypes.client
import win32com.client

from fpdf import FPDF
from pytesseract import pytesseract
from PyPDF2 import PdfFileMerger


def convert_to_pdf(input_file, output_file):
    if isinstance(input_file, str):
        if os.path.splitext(input_file)[1] in ['.doc', '.docx']:
            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(input_file)
            doc.SaveAs(output_file, FileFormat=17)
            doc.Close()
            word.Quit()

            return True
        elif os.path.splitext(input_file)[1] == '.txt':
            pdf = _set_pdf_settings()
            file = open(input_file, "r")
            for word in file:
                pdf.multi_cell(170, 10, txt=word, align='L')

            pdf.output(output_file)
            return True
        elif os.path.splitext(input_file)[1].lower() in ['.png', '.jpg', '.jpeg']:
            custom_config = '-l eng --oem 1 --psm 3'
            # Replace here the tesseract executable location to pytesseract library
            # For most installations the path would be C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe.
            pytesseract.tesseract_cmd = 'path_to_tesseract_exe'
            pdf = _set_pdf_settings()
            img = cv2.imread(input_file)
            text = pytesseract.image_to_string(img, config=custom_config)
            pdf.multi_cell(170, 10, txt=text, align='L')
            pdf.output(output_file)

            return True
    else:
        return False


def _set_pdf_settings():
    pdf = FPDF()
    pdf.set_font("Arial", size=11)
    pdf.set_left_margin(15)
    pdf.set_top_margin(15)
    pdf.set_right_margin(15)
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    return pdf


def convert_to_doc(input_file, output_file):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        worker = word.Documents.Open(input_file)
        worker.SaveAs2(output_file, FileFormat=16)
        worker.Close()

        word.Quit()

        return True
    except Exception as error:
        return str(error)


def merge_pdfs(pdf_list, output_pdf_name):
    merger = PdfFileMerger(strict=False)
    try:
        for pdf in pdf_list:
            merger.append(pdf)

        pdf_output_file = open(output_pdf_name, 'wb')
        merger.write(pdf_output_file)
        merger.close()

        return True

    except Exception as error:
        return str(error)
