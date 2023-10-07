import ctypes
import locale
import sys
from PyQt5.QtWidgets import *
from modules.MainWindow import MainWindow


myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    locale.setlocale(locale.LC_ALL, 'ru_RU.utf8')

    window = MainWindow()

    window.show()
    sys.exit(app.exec_())

