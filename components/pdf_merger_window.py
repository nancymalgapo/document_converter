import os
import sys

from PyQt5.QtWidgets import (QDialog, QFileDialog, QMessageBox)

from components.logic_methods import merge_pdfs
from ui.pdf_merger import Ui_PDFMerger


class PDFMergerWindow(QDialog, Ui_PDFMerger):
    def __init__(self):
        super().__init__()
        self.pdf_merger_ui = Ui_PDFMerger()
        self.pdf_merger_ui.setupUi(self)
        self._input_pdf_files = list()
        self._output_pdf_file = None
        self.connect_signals_and_slots()

    def connect_signals_and_slots(self):
        self.pdf_merger_ui.input_browse_button.clicked.connect(self._select_pdf_files)
        self.pdf_merger_ui.output_browse_button.clicked.connect(self._choose_output_filename)
        self.pdf_merger_ui.proceed_button.clicked.connect(self._proceed_to_operation)

    def _select_pdf_files(self):
        file_path = QFileDialog.getOpenFileNames(self, 'Select Files', '', 'PDF File (*.pdf)')[0]
        if isinstance(file_path, list):
            self._input_pdf_files = file_path
            self.pdf_merger_ui.input_line_edit.setText(str(file_path).replace("'", "")[1:-1])

    def _choose_output_filename(self):
        self._output_pdf_file, check = QFileDialog.getSaveFileName(self, 'Save File', '', 'PDF File (*.pdf)')
        if check:
            ext_file = os.path.splitext(self._output_pdf_file)[1]
            self.pdf_merger_ui.output_line_edit.setText("{0}.pdf".format(self._output_pdf_file)
                                                        if ext_file != '.pdf' else self._output_pdf_file)
            self._output_pdf_file = self.pdf_merger_ui.output_line_edit.text()

    def _proceed_to_operation(self):
        result = merge_pdfs(self._input_pdf_files, self._output_pdf_file)
        message_box = QMessageBox()
        message_box.setStandardButtons(QMessageBox.Ok)
        if result is True:
            message_box.setIcon(QMessageBox.Information)
            message_box.setText("Operation Success")
            message_box.setWindowTitle("PDF Files are successfully merged.")
            message_box.exec()
            sys.exit()
        else:
            message_box.setIcon(QMessageBox.Warning)
            message_box.setText("Operation Failed")
            message_box.setWindowTitle("Merging of PDF files failed. Please check and try again.")
            return_value = message_box.exec()
            if return_value == QMessageBox.Ok:
                message_box.close()
