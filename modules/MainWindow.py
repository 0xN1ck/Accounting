from widgets.Ui_MainWindow import Ui_MainWindow
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from utils.garantex import garantex_history
from utils.bybit import bybit_history
from utils.exnode import exnode_history
from utils.huobi import huobi_history
from utils.excel.main_file import create_new_month
# from utils.test import test
import sys


class MainWindow(QMainWindow):
    """
        Класс MainWindow является основным окном приложения. Он наследуется от класса QMainWindow и содержит графический
    интерфейс, созданный с помощью файла widgets.Ui_MainWindow.

        В конструкторе класса инициализируются элементы интерфейса, устанавливаются обработчики событий и привязываются
    функции к кнопкам.

        Метод update_pb_ax_all обновляет состояние кнопки pb_ex_all в зависимости от того, заполнены ли все поля ввода.

        Метод execute_all выполняет последовательное выполнение функций garantex_history, bybit_history, exnode_history
    и huobi_history. Если выполнение хотя бы одной из функций вернет 0, то выполнение остальных функций прерывается.

        Метод create_main_file вызывает функцию create_new_month для создания основного файла.

        Метод handle_global_exception перехватывает и обрабатывает глобальные исключения, выводя сообщение об ошибке
    в диалоговом окне и в статусной строке.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        sys.excepthook = self.handle_global_exception

        self.path_all = {
            'file': '',
            'garantex': '',
            'bybit': '',
            'exnode': '',
            'huobi': ''
        }

        self.label_status = QLabel(self)
        self.ui.statusbar.addWidget(self.label_status)

        self.ui.action_redact_config.triggered.connect(self.ui.cb_user.open_config_file)
        self.ui.action_create_month.triggered.connect(self.create_main_file)

        self.ui.le_file.textChanged.connect(self.update_pb_ax_all)
        self.ui.le_garantex.textChanged.connect(self.update_pb_ax_all)
        self.ui.le_bybit.textChanged.connect(self.update_pb_ax_all)
        self.ui.le_exnode.textChanged.connect(self.update_pb_ax_all)
        self.ui.le_huobi.textChanged.connect(self.update_pb_ax_all)
        self.ui.le_all_market.textChanged.connect(self.update_pb_ax_all)

        self.ui.pb_ex_garantex.clicked.connect(lambda: garantex_history(self))
        self.ui.pb_ex_bybit.clicked.connect(lambda: bybit_history(self))
        self.ui.pb_ex_exnode.clicked.connect(lambda: exnode_history(self))
        self.ui.pb_ex_huobi.clicked.connect(lambda: huobi_history(self))
        self.ui.pb_ex_all.clicked.connect(self.execute_all)

        current_datetime = QDateTime.currentDateTime()
        self.ui.date_start.setDateTime(current_datetime.addMSecs(-current_datetime.time().minute() * 60 * 1000))
        self.ui.date_finish.setDateTime(
            current_datetime.addDays(1).addMSecs(-current_datetime.time().minute() * 60 * 1000))
        self.ui.pb_ex_all.setEnabled(False)

        # test(self)

    def update_pb_ax_all(self):
        if (
                (self.ui.le_file.text() and
                 self.ui.le_garantex.text() and
                 self.ui.le_bybit.text() and
                 self.ui.le_exnode.text() and
                 self.ui.le_huobi.text()) or
                self.ui.le_all_market.text()
        ):
            self.ui.pb_ex_all.setEnabled(True)
        else:
            self.ui.pb_ex_all.setEnabled(False)

    def execute_all(self):
        funcs = [garantex_history, bybit_history, exnode_history, huobi_history]
        for func in funcs:
            if func(self) == 0:
                return
        self.label_status.setText(f'<font color="green">Расчет бирж выполнен для пользователя '
                                  f'{self.ui.cb_user.currentText()}</font>')

    def create_main_file(self):
        create_new_month(self)

    def handle_global_exception(self, exc_type, exc_value, exc_traceback):
        error_message = f"Произошла ошибка: {exc_value}"
        error_box = QMessageBox(QMessageBox.Critical, "Ошибка", error_message, QMessageBox.Ok, self)
        error_box.show()
        self.label_status.setText(f'<font color="red">Произошла ошибка: {exc_value}</font>')
