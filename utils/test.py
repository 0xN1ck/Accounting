from PyQt5.QtCore import *
# import os


def test(self):
    self.ui.date_start.setDateTime(QDateTime(QDate(2023, 9, 11), QTime(19, 0, 0)))
    self.ui.date_finish.setDateTime(QDateTime(QDate(2023, 9, 12), QTime(7, 0, 0)))

    self.ui.le_file.setText("Сентябрь.xlsx")
    self.ui.le_garantex.setText("garantex")
    self.ui.le_bybit.setText("bybit")
    self.ui.le_exnode.setText("exnode")
    self.ui.le_huobi.setText("huobi")

    self.path_all = {
        'file': 'docs/Сентябрь.xlsx',
        'garantex': 'docs/garantex',
        'bybit': 'docs/bybit',
        'exnode': 'docs/exnode',
        'huobi': 'docs/huobi'
    }