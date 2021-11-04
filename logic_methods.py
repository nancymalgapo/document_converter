import os
import comtypes.client

from fpdf import FPDF
from PyPDF2 import PdfFileMerger


def convert_to_pdf(input_file, output_file):
    if os.path.splitext(input_file)[1] in ['.doc', '.docx']:
        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(input_file)
        doc.SaveAs(output_file, FileFormat=17)
        doc.Close()
        word.Quit()

        return True
    elif os.path.splitext(input_file)[1] == '.txt':
        pdf = FPDF()
        pdf.set_font("Arial", size=11)
        pdf.set_left_margin(15)
        pdf.set_top_margin(15)
        pdf.set_right_margin(15)
        pdf.set_auto_page_break(True, margin=15)
        pdf.add_page()
        file = open(input_file, "r")
        for word in file:
            pdf.multi_cell(170, 10, txt=word, align='L')

        pdf.output(output_file)
        return True
    else:
        return False


def pdf_to_word(input_file, output_file):
    print('Done converting {0} to {1}'.format(input_file, output_file))


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