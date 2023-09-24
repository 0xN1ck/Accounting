import pandas as pd
from utils.excel import *


def exnode_history(self):
    self.current_sheet = check_current_result_file(self)
    if not self.current_sheet:
        self.label_status.setText(f'<font color="red">Не найден лист {self.ui.cb_user.currentText()}'
                                  f'({self.ui.date_start.dateTime().toString("d HH-mm")}-'
                                  f'{self.ui.date_finish.dateTime().toString("d HH-mm")}). </font>'
                                  f'<font color="green">Активируйте создание листа</font>')
        return

    exnode_history = pd.read_excel(self.path_all['exnode'] + '/exnode история.xlsx')

    date_start = self.ui.date_start.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")
    date_finish = self.ui.date_finish.dateTime().toPyDateTime().strftime("%Y-%m-%d %H:%M:%S")
    exnode_history['create_time'] = pd.to_datetime(exnode_history['create_time']).dt.strftime("%Y-%m-%d %H:%M:%S")
    filtered_data_exnode_history = exnode_history[(exnode_history['create_time'] >= f'{date_start}') & (
            exnode_history['create_time'] <= f'{date_finish}')]

    write_to_excel_exnode_history(self.current_sheet['filename'],
                                  self.current_sheet['current_sheet'],
                                  self.current_sheet['workbook'],
                                  filtered_data_exnode_history)

    write_to_excel_exnode_komsa(self.current_sheet['filename'],
                                self.current_sheet['current_sheet'],
                                self.current_sheet['workbook'],
                                filtered_data_exnode_history)

    self.label_status.setText('<font color="green">EXNODE выполнен</font>')
