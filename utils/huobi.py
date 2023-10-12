from datetime import timedelta
from utils.excel.huobi import *


def huobi_history(self):
    """
        Функция `huobi_history` используется для обработки и записи данных из файла истории P2P-ордеров биржи Huobi
    в Excel-файл.

        Сначала функция проверяет наличие текущего листа в Excel-файле результатов. Если текущий лист не найден,
    выводится сообщение об ошибке и функция завершается.

        Затем функция загружает данные из файла истории P2P-ордеров биржи Huobi в DataFrame и преобразует столбец
    с временем в формат datetime.

        Затем функция получает даты начала и конца периода, выбранные пользователем в графическом интерфейсе,
    и фильтрует данные по этому периоду.

        Затем функция фильтрует данные по типу операции (покупка или продажа) и записывает отфильтрованные данные
    в Excel-файл результатов.

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

    huobi_history_p2p = pd.read_excel(self.path_all['huobi'] + '/huobi история p2p-ордеров.xlsx')

    huobi_history_p2p['Время'] = pd.to_datetime(huobi_history_p2p['Время'])

    date_start = (self.ui.date_start.dateTime().toPyDateTime() + timedelta(hours=5)).strftime("%Y-%m-%d %H:%M:%S")
    date_finish = (self.ui.date_finish.dateTime().toPyDateTime() + timedelta(hours=5)).strftime("%Y-%m-%d %H:%M:%S")
    filtered_data_huobi_history_p2p = huobi_history_p2p[(huobi_history_p2p['Время'] >= f'{date_start}') &
                                                        (huobi_history_p2p['Время'] <= f'{date_finish}')]

    filtered_data_huobi_history_p2p_buy = filtered_data_huobi_history_p2p[
        filtered_data_huobi_history_p2p['Тип'] == 'Купить']

    write_to_to_excel_huobi_history_p2p_buy(self.current_sheet['filename'], self.current_sheet['current_sheet'],
                                            self.current_sheet['workbook'], filtered_data_huobi_history_p2p_buy)

    filtered_data_huobi_history_p2p_sell = filtered_data_huobi_history_p2p[
        filtered_data_huobi_history_p2p['Тип'] == 'Продать']

    write_to_to_excel_huobi_history_p2p_sell(self.current_sheet['filename'], self.current_sheet['current_sheet'],
                                             self.current_sheet['workbook'], filtered_data_huobi_history_p2p_sell)

    self.label_status.setText('<font color="green">huobi выполнен</font>')
