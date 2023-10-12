from datetime import timedelta
from utils.excel.bybit import *


def bybit_history(self):
    """
        Функция `bybit_history` используется для обработки и записи истории сделок на бирже ByBit в Excel-файл.

        Сначала функция проверяет наличие текущего листа в Excel-файле результатов. Если лист не найден, выводится
    сообщение об ошибке и функция завершается.

        Затем функция загружает данные истории сделок на бирже ByBit из файла "bybit история p2p сделок.xlsx"
    и преобразует столбец с датой и временем в формат datetime.

        Далее функция определяет дату начала и конца периода, указанного пользователем, и фильтрует данные истории
    сделок по этому периоду.

        Затем функция разделяет отфильтрованные данные на две части: сделки на покупку (BUY) и сделки на продажу (SELL).

        Функция записывает отфильтрованные данные о сделках на покупку в текущий лист Excel-файла результатов,
    используя функцию `write_to_excel_bybit_history_p2p`.

        Затем функция записывает отфильтрованные данные о сделках на продажу в текущий лист Excel-файла результатов,
    используя функцию `write_to_excel_bybit_history_p2p_sell`.

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

    bybit_history_p2p = pd.read_excel(self.path_all['bybit'] + '/bybit история p2p сделок.xlsx')

    bybit_history_p2p['Time'] = pd.to_datetime(bybit_history_p2p['Time'], format='%Y-%m-%d %H:%M:%S', dayfirst=True)
    date_start = (self.ui.date_start.dateTime().toPyDateTime() - timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
    date_finish = (self.ui.date_finish.dateTime().toPyDateTime() - timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")

    filtered_data_bybit_history_p2p = bybit_history_p2p[(bybit_history_p2p['Time'] >= f'{date_start}') &
                                                        (bybit_history_p2p['Time'] <= f'{date_finish}')]

    filtered_data_bybit_history_p2p_buy = filtered_data_bybit_history_p2p[
        filtered_data_bybit_history_p2p['Type'] == 'BUY']
    filtered_data_bybit_history_p2p_buy.set_index('Cryptocurrency', inplace=True)

    filtered_data_bybit_history_p2p_sell = filtered_data_bybit_history_p2p[
        filtered_data_bybit_history_p2p['Type'] == 'SELL']

    write_to_excel_bybit_history_p2p(self.current_sheet['filename'],
                                     self.current_sheet['current_sheet'],
                                     self.current_sheet['workbook'],
                                     filtered_data_bybit_history_p2p_buy)

    write_to_excel_bybit_history_p2p_sell(self.current_sheet['filename'],
                                          self.current_sheet['current_sheet'],
                                          self.current_sheet['workbook'],
                                          filtered_data_bybit_history_p2p_sell)

    self.label_status.setText('<font color="green">ByBit выполнен</font>')
