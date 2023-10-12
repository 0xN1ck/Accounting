from PyQt5.QtCore import *
# import os


def test(self):
    """
        Функция "test" устанавливает значения даты и времени в виджетах "date_start" и "date_finish" интерфейса
    пользователя. Значения устанавливаются на 28 сентября 2023 года, со временем 9:00 и 19:10 соответственно.

        Комментарии к строкам, которые закомментированы, указывают на то, что они были закомментированы и, возможно,
    ранее использовались для установки значений в других виджетах или переменных.

        В конце функции создается словарь "path_all", который содержит пути к различным файлам и директориям.
    Ключи словаря соответствуют различным объектам, таким как "file", "garantex", "bybit", "exnode" и "huobi".
    Значения путей указаны относительно директории "docs".

    :param self:
    :return:
    """
    self.ui.date_start.setDateTime(QDateTime(QDate(2023, 9, 28), QTime(9, 0, 0)))
    self.ui.date_finish.setDateTime(QDateTime(QDate(2023, 9, 28), QTime(19, 10, 0)))

    # self.ui.le_file.setText("Сентябрь.xlsx")
    # self.ui.le_garantex.setText("garantex")
    # self.ui.le_bybit.setText("bybit")
    # self.ui.le_exnode.setText("exnode")
    # self.ui.le_huobi.setText("huobi")

    # self.path_all = {
    #     'file': 'docs/Сентябрь 2023.xlsx',
    #     'garantex': 'docs/garantex',
    #     'bybit': 'docs/bybit',
    #     'exnode': 'docs/exnode',
    #     'huobi': 'docs/huobi'
    # }
