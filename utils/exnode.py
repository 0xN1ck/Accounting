import pandas as pd
from utils.excel.exnode import *


def exnode_history(self):
    """
        Функция `exnode_history` отвечает за создание истории операций в системе EXNODE.

        Сначала функция проверяет наличие текущего листа в результатном файле. Если лист не найден, выводится сообщение
    об ошибке и предлагается активировать создание листа.

        Затем функция загружает данные из файлов "orders.xlsx" и "вводы_выводы.xlsx" в формате Excel с помощью
    библиотеки pandas. Даты начала и конца периода выбираются пользователем с помощью виджетов QDateTimeEdit
    и преобразуются в строковый формат.

        Далее происходит фильтрация данных по заданному периоду. Фильтрованные данные записываются в текущий лист
    в главном файле с помощью функций `write_to_excel_exnode_orders` и `write_to_excel_exnode_komsa`.

        В конце функция выводит сообщение об успешном выполнении операции.

    :param self:
    :return:
    """
    self.current_sheet = check_current_result_file(self)
    if not self.current_sheet:
        self.label_status.setText(f'<font color="red">Не найден лист {self.ui.cb_user.currentText()}'
                                  f'({self.ui.date_start.dateTime().toString("d HH-mm")}-'
                                  f'{self.ui.date_finish.dateTime().toString("d HH-mm")}). </font>'
                                  f'<font color="green">Активируйте создание листа</font>')
        return 0

    exnode_orders = pd.read_excel(self.path_all['exnode'] + '/orders.xlsx')

    date_start = self.ui.date_start.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")
    date_finish = self.ui.date_finish.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")

    exnode_orders['create_time'] = pd.to_datetime(exnode_orders['create_time']).dt.strftime("%Y-%m-%d %H:%M:%S")
    filtered_data_exnode_orders = exnode_orders[(exnode_orders['create_time'] >= f'{date_start}') & (
            exnode_orders['create_time'] <= f'{date_finish}')]

    exnode_in_out = pd.read_excel(self.path_all['exnode'] + '/вводы_выводы.xlsx')

    filtered_data_exnode_in_out = exnode_in_out[(exnode_in_out['date_create'] >= f'{date_start}') & (
            exnode_in_out['date_create'] <= f'{date_finish}')]

    write_to_excel_exnode_orders(self.current_sheet['filename'],
                                 self.current_sheet['current_sheet'],
                                 self.current_sheet['workbook'],
                                 filtered_data_exnode_orders)
    #
    write_to_excel_exnode_komsa(self.current_sheet['filename'],
                                self.current_sheet['current_sheet'],
                                self.current_sheet['workbook'],
                                filtered_data_exnode_in_out)

    self.label_status.setText('<font color="green">EXNODE выполнен</font>')
