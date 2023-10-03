import sys, locale, ctypes
# from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from modules.MainWindow import MainWindow


myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

app = QApplication(sys.argv)
locale.setlocale(locale.LC_ALL, 'ru_RU.utf8')

window = MainWindow()

window.show()
sys.exit(app.exec_())

