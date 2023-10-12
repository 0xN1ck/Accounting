from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class LineEditFilePath(QLineEdit):
    """
        Класс `LineEditFilePath` наследуется от класса `QLineEdit` и представляет собой текстовое поле для
    отображения пути к файлу.

    В конструкторе класса устанавливаются начальные значения атрибутов и устанавливается режим только для чтения.

    Метод `open_file_dialog` вызывается при нажатии на поле мышью и открывает диалоговое окно для выбора файла.

        В диалоговом окне устанавливается режим выбора только существующих файлов и фильтр для отображения только файлов
    Excel (.xlsx и .xls).

        Если пользователь выбирает файл, то его путь сохраняется в атрибуте `file_path`, из которого извлекается
    имя файла.

        Затем путь к файлу сохраняется в словаре `path_all` родительского виджета и отображается в текстовом поле.
    """
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
            self.parent().parent().path_all['file'] = self.file_path
            self.setText(file_name)