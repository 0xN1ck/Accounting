from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class LineEditMarketPath(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.market_path = ""

        self.setReadOnly(True)
        self.mousePressEvent = self.open_folder_dialog

    def open_folder_dialog(self, event):
        folder_dialog = QFileDialog()
        folder_dialog.setFileMode(QFileDialog.Directory)
        if folder_dialog.exec_():
            self.market_path = folder_dialog.selectedFiles()[0]
            market_info = QFileInfo(self.market_path)
            market_name = market_info.fileName()
            self.setText(market_name)