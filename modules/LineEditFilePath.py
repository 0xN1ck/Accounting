from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class LineEditFilePath(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.file_path = ""

        self.setReadOnly(True)
        self.mousePressEvent = self.open_file_dialog

    def open_file_dialog(self, event):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        if file_dialog.exec_():
            self.file_path = file_dialog.selectedFiles()[0]
            file_info = QFileInfo(self.file_path)
            file_name = file_info.fileName()
            self.setText(file_name)