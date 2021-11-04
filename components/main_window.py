from PyQt5.QtWidgets import QMainWindow

from components.converter_window import ConvertToDocWindow, ConvertToPDFWindow
from components.pdf_merger_window import PDFMergerWindow
from ui.app_main_window import Ui_MainWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self._main_ui = Ui_MainWindow()
        self._main_ui.setupUi(self)
        self._word_to_pdf_window = ConvertToPDFWindow()
        self._pdf_to_word_window = ConvertToDocWindow()
        self._pdf_merger_window = PDFMergerWindow()
        self._connect_signals_and_slots()

    def _connect_signals_and_slots(self):
        self._main_ui.proceed_button.clicked.connect(self._open_secondary_window)

    def _open_secondary_window(self):
        if self._main_ui.word_to_pdf_option.isChecked():
            self._word_to_pdf_window.show()
        elif self._main_ui.pdf_to_word_button.isChecked():
            self._pdf_to_word_window.show()
        else:
            self._pdf_merger_window.show()
