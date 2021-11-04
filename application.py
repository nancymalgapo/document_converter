import sys

from PyQt5.QtWidgets import (QDialog, QMainWindow, QFileDialog, QMessageBox)

from logic_methods import *
from ui.app_main_window import Ui_MainWindow
from ui.word_to_pdf import Ui_WordToPDF
from ui.pdf_merger import Ui_PDFMerger


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.main_ui = Ui_MainWindow()
        self.main_ui.setupUi(self)
        self.converter_window = ConverterWindow()
        self.pdf_merger_window = PDFMergerWindow()
        self.connect_signals_and_slots()

    def connect_signals_and_slots(self):
        self.main_ui.proceed_button.clicked.connect(self._open_secondary_window)

    def _open_secondary_window(self):
        if self.main_ui.word_to_pdf_option.isChecked():
            self.converter_window.show()
        else:
            self.pdf_merger_window.show()


class ConverterWindow(QDialog, Ui_WordToPDF):
    def __init__(self):
        super().__init__()
        self.converter_ui = Ui_WordToPDF()
        self.converter_ui.setupUi(self)
        self._input_file = None
        self._output_file = None
        self.connect_signals_and_slots()

    def connect_signals_and_slots(self):
        self.converter_ui.browse_input_button.clicked.connect(self._open_file_dialog)
        self.converter_ui.browse_output_button.clicked.connect(self._save_file_to_location)
        self.converter_ui.convert_button.clicked.connect(self._show_converter_dialog)

    def _open_file_dialog(self):
        self._input_file, check = QFileDialog.getOpenFileName(self, 'Open File', '',
                                                              'Doc Files (*.doc);;Docx Files (*.docx);;'
                                                              'Text Files (*.txt)')
        if check:
            self.converter_ui.input_line_edit.setText(self._input_file)

    def _save_file_to_location(self):
        self._output_file, check = QFileDialog.getSaveFileName(self, 'Save File', '', 'PDF File (*.pdf)')
        if check:
            ext_file = os.path.splitext(self._output_file)[1]
            self.converter_ui.output_line_edit.setText("{0}.pdf".format(self._output_file)
                                                       if ext_file != '.pdf' else self._output_file)
            self._output_file = self.converter_ui.output_line_edit.text()

    def _show_converter_dialog(self):
        process = convert_to_pdf(self._input_file, self._output_file)
        message_box = QMessageBox()
        message_box.setStandardButtons(QMessageBox.Ok)
        if process:
            message_box.setIcon(QMessageBox.Information)
            message_box.setText("Operation Success")
            message_box.setWindowTitle("File is successfully converted.")
            message_box.exec()
            sys.exit()
        else:
            message_box.setIcon(QMessageBox.Warning)
            message_box.setText("Operation Failed")
            message_box.setWindowTitle("File conversion failed. Please check and try again.")
            return_value = message_box.exec()
            if return_value == QMessageBox.Ok:
                message_box.close()


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
