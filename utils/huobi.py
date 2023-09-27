import pandas as pd
from datetime import timedelta
from utils.excel import *


def huobi_history(self):
    self.current_sheet = check_current_result_file(self)
    if not self.current_sheet:
        self.label_status.setText(f'<font color="red">Не найден лист {self.ui.cb_user.currentText()}'
                                  f'({self.ui.date_start.dateTime().toString("d HH-mm")}-'
                                  f'{self.ui.date_finish.dateTime().toString("d HH-mm")}). </font>'
                                  f'<font color="green">Активируйте создание листа</font>')
        return

    huobi_history_p2p = pd.read_excel(self.path_all['huobi'] + '/huobi история p2p-ордеров.xlsx')

    huobi_history_p2p['Время'] = pd.to_datetime(huobi_history_p2p['Время'])
    huobi_history_p2p['Время'] = huobi_history_p2p['Время'] - timedelta(hours=5)

    date_start = self.ui.date_start.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")
    date_finish = self.ui.date_finish.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")
    filtered_data_huobi_history_p2p = huobi_history_p2p[(huobi_history_p2p['Время'] >= f'{date_start}') &
                                                        (huobi_history_p2p['Время'] <= f'{date_finish}')]

    filtered_data_huobi_history_p2p_buy = filtered_data_huobi_history_p2p[
        filtered_data_huobi_history_p2p['Тип'] == 'Купить']
    huobi_history_p2p_buy_sum_by_coin = filtered_data_huobi_history_p2p_buy.groupby('Монета').agg(
        {'Количество': 'sum', 'Общая цена': 'sum'})

    huobi_history_spot = pd.read_csv(self.path_all['huobi'] + '/huobi история спотовой торговли.csv')

    huobi_history_spot['Дата'] = pd.to_datetime(huobi_history_spot['Дата'])
    huobi_history_spot['Дата'] = huobi_history_spot['Дата'] - timedelta(hours=5)

    tolerance = timedelta(minutes=5)
    huobi_history_spot['Дата'] = pd.to_datetime(huobi_history_spot['Дата'])
    filtered_data_huobi_history_p2p_copy = filtered_data_huobi_history_p2p.copy()
    filtered_data_huobi_history_p2p_copy['Время'] = pd.to_datetime(filtered_data_huobi_history_p2p_copy['Время'])
    filtered_time_huobi_spot = huobi_history_spot[huobi_history_spot['Дата'].apply(
        lambda x: any(abs(x - filtered_data_huobi_history_p2p_copy['Время']) < tolerance))]

    filtered_time_huobi_spot = filtered_time_huobi_spot.copy()
    filtered_time_huobi_spot['Комиссия'] = filtered_time_huobi_spot['Комиссия'].astype(str).str.replace(r'[^\d.]', '',
                                                                                                        regex=True)
    filtered_time_huobi_spot['Комиссия'] = pd.to_numeric(filtered_time_huobi_spot['Комиссия'])

    huobi_spot_mean_sum = filtered_time_huobi_spot.groupby('Пара').agg({'Цена': 'mean', 'Комиссия': 'sum'})
    huobi_spot_mean_sum.reset_index(inplace=True)

    write_to_to_excel_huobi_history_p2p_buy(self.current_sheet['filename'], self.current_sheet['current_sheet'],
                                            self.current_sheet['workbook'],
                                            huobi_history_p2p_buy_sum_by_coin,
                                            huobi_spot_mean_sum)

    write_to_excel_huobi_komsa(self.current_sheet['filename'], self.current_sheet['current_sheet'],
                               self.current_sheet['workbook'], huobi_spot_mean_sum)

    self.label_status.setText('<font color="green">huobi выполнен</font>')
