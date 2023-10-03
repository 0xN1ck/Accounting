from widgets.Ui_MainWindow import Ui_MainWindow
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from utils.garantex import garantex_history
from utils.bybit import bybit_history
from utils.exnode import exnode_history
from utils.huobi import huobi_history
from utils.excel import create_new_month
# from utils.test import test


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

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

        # print(self.ui.lb_garantex.text().lower())
        # test(self)

    def update_pb_ax_all(self):
        if (
                self.ui.le_file.text() and
                self.ui.le_garantex.text() and
                self.ui.le_bybit.text() and
                self.ui.le_exnode.text() and
                self.ui.le_huobi.text()
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
