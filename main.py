import sys, locale, ctypes
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from modules.MainWindow import MainWindow

myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

app = QApplication(sys.argv)
# Установка русской локали
locale.setlocale(locale.LC_ALL, 'ru_RU.utf8')


# style_sheet_file = QFile("resources/sheets/QSS-Stylesheets-master/MacOS.qss")
# style_sheet_file.open(QFile.ReadOnly)
# style_sheet = style_sheet_file.readAll()
# app.setStyleSheet(style_sheet.data().decode("UTF-8"))

window = MainWindow()

window.show()
sys.exit(app.exec_())

