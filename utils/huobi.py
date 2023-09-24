import pandas as pd
from datetime import timedelta
from utils.excel import *


def exnode_history(self):
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

