import pandas as pd
from datetime import timedelta
from utils.excel import *


def bybit_history(self):
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

    filtered_data_bybit_history_p2p_buy_sum = filtered_data_bybit_history_p2p_buy.groupby('Cryptocurrency').agg(
        {'Fiat Amount': 'sum', 'Price': 'mean', 'Coin Amount': 'sum'})

    bybit_history_spot = pd.read_excel(self.path_all['bybit'] + '/bybit история спотовой торговли.xlsx')

    bybit_history_spot['Timestamp (Local Time)'] = pd.to_datetime(bybit_history_spot['Timestamp (Local Time)'],
                                                                  format='%Y-%m-%d %H:%M:%S', dayfirst=True)

    filtered_data_bybit_history_spot = bybit_history_spot[
        (bybit_history_spot[
             'Timestamp (Local Time)'] >= f'{self.ui.date_start.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")}') &
        (bybit_history_spot[
             'Timestamp (Local Time)'] <= f'{self.ui.date_finish.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")}')
        ]

    for index in filtered_data_bybit_history_p2p_buy_sum.index:
        if not index == 'USDT':
            matching_values = \
                filtered_data_bybit_history_spot[filtered_data_bybit_history_spot['Spot Pairs'].str.contains(index)][
                    'Filled Value'].apply(lambda x: re.findall(r'\d+\.\d+', x)[0]).astype(float).sum()
            filtered_data_bybit_history_p2p_buy_sum.loc[index]['Coin Amount'] = matching_values
            filtered_data_bybit_history_p2p_buy_sum.loc[index]['Price'] = \
                filtered_data_bybit_history_p2p_buy_sum.loc[index]['Fiat Amount'] / matching_values
            matching_values = 0

    start_index = filtered_data_bybit_history_spot.index.values[0]
    end_index = filtered_data_bybit_history_spot.index.values[-1] + 3

    filtered_data_bybit_history_spot = bybit_history_spot.iloc[start_index:end_index]

    filtered_data_bybit_history_spot_copy = filtered_data_bybit_history_spot.copy()

    filtered_data_bybit_history_spot_copy.loc[:, 'Spot Pairs'] = pd.to_datetime(
        filtered_data_bybit_history_spot_copy['Spot Pairs'], errors='coerce')

    filtered_data_bybit_for_komsa = filtered_data_bybit_history_spot_copy[
        pd.notnull(filtered_data_bybit_history_spot_copy['Spot Pairs'])]

    write_to_excel_bybit_history_p2p(self.current_sheet['filename'],
                                     self.current_sheet['current_sheet'],
                                     self.current_sheet['workbook'],
                                     filtered_data_bybit_history_p2p_buy_sum)

    write_to_excel_bybit_komsa(self.current_sheet['filename'],
                               self.current_sheet['current_sheet'],
                               self.current_sheet['workbook'],
                               filtered_data_bybit_for_komsa,
                               filtered_data_bybit_history_p2p_buy_sum)

    filtered_data_bybit_history_p2p_sell = filtered_data_bybit_history_p2p[
        filtered_data_bybit_history_p2p['Type'] == 'SELL']

    write_to_excel_bybit_history_p2p_sell(self.current_sheet['filename'],
                                          self.current_sheet['current_sheet'],
                                          self.current_sheet['workbook'],
                                          filtered_data_bybit_history_p2p_sell)

    self.label_status.setText('<font color="green">ByBit выполнен</font>')
