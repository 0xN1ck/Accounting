import pandas as pd
from utils.excel.garantex import *


def garantex_history(self):
    self.current_sheet = check_current_result_file(self)
    if not self.current_sheet:
        self.label_status.setText(f'<font color="red">Не найден лист {self.ui.cb_user.currentText()}'
                                  f'({self.ui.date_start.dateTime().toString("d HH-mm")}-'
                                  f'{self.ui.date_finish.dateTime().toString("d HH-mm")}). </font>'
                                  f'<font color="green">Активируйте создание листа</font>')
        return 0

    gara_history_p2p = pd.read_excel(self.path_all['garantex'] + '/garantex история p2p сделок.xlsx')
    gara_history_p2p['Дата и время'] = pd.to_datetime(gara_history_p2p['Дата и время'], dayfirst=True)

    date_start = self.ui.date_start.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")
    date_finish = self.ui.date_finish.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")

    filtered_data_gara_history_p2p = gara_history_p2p[
        (gara_history_p2p['Дата и время'] >= f'{date_start}') & (gara_history_p2p['Дата и время'] <= f'{date_finish}')
        ]

    pd.options.mode.chained_assignment = None
    filtered_data_gara_history_p2p.loc[filtered_data_gara_history_p2p['Направление'] == 'Покупка', 'Результат'] = \
        filtered_data_gara_history_p2p['Итого'] - filtered_data_gara_history_p2p['Fiat']
    filtered_data_gara_history_p2p.loc[filtered_data_gara_history_p2p['Направление'] == 'Продажа', 'Результат'] = \
        filtered_data_gara_history_p2p['Итого'] + filtered_data_gara_history_p2p['Fiat']

    write_to_excel_gara_history_p2p(self.current_sheet['filename'],
                                    self.current_sheet['current_sheet'],
                                    self.current_sheet['workbook'],
                                    filtered_data_gara_history_p2p)

    gara_history_swap = pd.read_excel(self.path_all['garantex'] + '/garantex история обменов.xlsx')

    gara_history_swap['Дата и время'] = pd.to_datetime(gara_history_swap['Дата и время'], dayfirst=True)
    filtered_data_gara_history_swap = gara_history_swap[(gara_history_swap['Дата и время'] >= f'{date_start}') & (
            gara_history_swap['Дата и время'] <= f'{date_finish}')]

    write_to_excel_gara_swap_history(self.current_sheet['filename'],
                                     self.current_sheet['current_sheet'],
                                     self.current_sheet['workbook'],
                                     filtered_data_gara_history_swap)

    gara_history_account = pd.read_excel(self.path_all['garantex'] + '/garantex история счета.xlsx',
                                         sheet_name='Withdraws')

    gara_history_account['Дата и время'] = pd.to_datetime(gara_history_account['Дата и время'], dayfirst=True)
    filtered_data_gara_history_account = gara_history_account[
        (gara_history_account['Дата и время'] >= f'{date_start}') &
        (gara_history_account['Дата и время'] <= f'{date_finish}')
        ]

    tolerance = pd.Timedelta(minutes=5)

    filtered_data_gara_history_account['Средняя цена'] = None

    for index, row in filtered_data_gara_history_account.iterrows():
        target_time = row['Дата и время']
        time_difference = abs(filtered_data_gara_history_swap['Дата и время'] - target_time)
        matching_row_index = time_difference.idxmin()
        if time_difference[matching_row_index] <= tolerance:
            filtered_data_gara_history_account.at[index, 'Средняя цена'] = filtered_data_gara_history_swap.at[
                matching_row_index, 'Средняя цена']

    write_to_excel_gara_account_history(self.current_sheet['filename'],
                                        self.current_sheet['current_sheet'],
                                        self.current_sheet['workbook'],
                                        filtered_data_gara_history_account)

    self.label_status.setText('<font color="green">Garantex выполнен</font>')
