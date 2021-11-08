import os
import sys

from PyQt5.QtWidgets import (QDialog, QFileDialog, QMessageBox)

from components.logic_methods import convert_to_doc, convert_to_pdf
from ui.word_to_pdf import Ui_WordToPDF
from ui.pdf_to_word import Ui_PDFToWord
from ui.image_to_pdf import Ui_ImageToPDF


class ParentWindowClass(QDialog):

    def __init__(self):
        super().__init__()
        self._input_file = None
        self._output_file = None

    def _connect_signals_and_slots(self):
        pass

    def _open_file_dialog(self):
        pass

    def _save_file_to_location(self):
        pass

    def _show_dialog(self):
        pass

    def _show_converter_dialog(self, procedure, success_msg, error_msg):
        process = procedure(self._input_file, self._output_file)
        message_box = QMessageBox()
        message_box.setStandardButtons(QMessageBox.Ok)
        if process:
            message_box.setIcon(QMessageBox.Information)
            message_box.setText("Operation Success")
            message_box.setWindowTitle(success_msg)
            message_box.exec()
            sys.exit()
        else:
            message_box.setIcon(QMessageBox.Warning)
            message_box.setText("Operation Failed")
            message_box.setWindowTitle(error_msg)
            return_value = message_box.exec()
            if return_value == QMessageBox.Ok:
                message_box.close()


class ConvertToPDFWindow(ParentWindowClass, Ui_WordToPDF):
    def __init__(self):
        super().__init__()
        self._converter_ui = Ui_WordToPDF()
        self._converter_ui.setupUi(self)
        self._connect_signals_and_slots()

    def _connect_signals_and_slots(self):
        self._converter_ui.browse_input_button.clicked.connect(self._open_file_dialog)
        self._converter_ui.browse_output_button.clicked.connect(self._save_file_to_location)
        self._converter_ui.convert_button.clicked.connect(self._show_dialog)

    def _open_file_dialog(self):
        self._input_file, check = QFileDialog.getOpenFileName(self, 'Open File', '',
                                                              'Doc Files (*.doc);;Docx Files (*.docx);;'
                                                              'Text Files (*.txt)')
        if check:
            self._converter_ui.input_line_edit.setText(self._input_file)

    def _save_file_to_location(self):
        self._output_file, check = QFileDialog.getSaveFileName(self, 'Save File', '', 'PDF File (*.pdf)')
        if check:
            ext_file = os.path.splitext(self._output_file)[1]
            self._converter_ui.output_line_edit.setText("{0}.pdf".format(self._output_file)
                                                        if ext_file != '.pdf' else self._output_file)
            self._output_file = self._converter_ui.output_line_edit.text()

    def _show_dialog(self):
        self._show_converter_dialog(convert_to_pdf, "File is successfully converted to PDF file.",
                                    "File conversion failed. Please check and try again.")


class ConvertToDocWindow(ParentWindowClass, Ui_PDFToWord):
    def __init__(self):
        super().__init__()
        self._converter_ui = Ui_PDFToWord()
        self._converter_ui.setupUi(self)
        self._connect_signals_and_slots()

    def _connect_signals_and_slots(self):
        self._converter_ui.browse_input_button.clicked.connect(self._open_file_dialog)
        self._converter_ui.browse_output_button.clicked.connect(self._save_file_to_location)
        self._converter_ui.convert_button.clicked.connect(self._show_dialog)

    def _open_file_dialog(self):
        self._input_file, check = QFileDialog.getOpenFileName(self, 'Open File', '', 'PDF Files (*.pdf)')
        if check:
            self._converter_ui.input_line_edit.setText(self._input_file)

    def _save_file_to_location(self):
        self._output_file, check = QFileDialog.getSaveFileName(self, 'Save File', '', 'Docx Files (*.docx)')
        if check:
            ext_file = os.path.splitext(self._output_file)[1]
            self._converter_ui.output_line_edit.setText("{0}.docx".format(self._output_file)
                                                        if ext_file != '.docx' else self._output_file)
            self._output_file = self._converter_ui.output_line_edit.text()

    def _show_dialog(self):
        self._show_converter_dialog(convert_to_doc, "File is successfully converted to Word Document.",
                                    "File conversion failed. Please check and try again.")


class ConvertFromImageToPDFWindow(ParentWindowClass, Ui_ImageToPDF):
    def __init__(self):
        super().__init__()
        self._converter_ui = Ui_ImageToPDF()
        self._converter_ui.setupUi(self)
        self._connect_signals_and_slots()

    def _connect_signals_and_slots(self):
        self._converter_ui.browse_input_button.clicked.connect(self._open_file_dialog)
        self._converter_ui.browse_output_button.clicked.connect(self._save_file_to_location)
        self._converter_ui.convert_button.clicked.connect(self._show_dialog)

    def _open_file_dialog(self):
        self._input_file, check = QFileDialog.getOpenFileNames(self, 'Select Images', '',
                                                               'PNG Files (*.png);;JPEG Files (*.jpeg);;'
                                                               'JPG Files (*.jpg)')
        if check:
            self._converter_ui.input_line_edit.setText(str(self._input_file)[1:-1])

    def _save_file_to_location(self):
        self._output_file, check = QFileDialog.getSaveFileName(self, 'Save File', '', 'PDF File (*.pdf)')
        if check:
            ext_file = os.path.splitext(self._output_file)[1]
            self._converter_ui.output_line_edit.setText("{0}.pdf".format(self._output_file)
                                                        if ext_file != '.pdf' else self._output_file)
            self._output_file = self._converter_ui.output_line_edit.text()

    def _show_dialog(self):
        self._show_converter_dialog(convert_to_pdf, "Successfully converted image/s to PDF File.",
                                    "File conversion failed. Please check and try again.")
