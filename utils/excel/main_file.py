import os
import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import coordinate_to_tuple


def create_new_month(self, flag=False):
    current_datetime = self.ui.date_start.dateTime().toPyDateTime()
    filename = './docs/' + current_datetime.strftime("%B %Y.xlsx")

    if os.path.exists(filename):
        self.label_status.setText(f'<font color="red">Файл уже создан </font>')
        return
    shutil.copyfile("docs/Шаблон.xlsx", filename)
    workbook = load_workbook(filename)
    template_sheet = workbook["Шаблон"]

    new_sheet = workbook.copy_worksheet(template_sheet)
    new_sheet.title = f'{self.ui.cb_user.currentText()}' \
                      f'({self.ui.date_start.dateTime().toString("d HH-mm")}-' \
                      f'{self.ui.date_finish.dateTime().toString("d HH-mm")})'
    workbook.active = len(workbook.sheetnames) - 1

    template_sheet.sheet_state = "hidden"

    workbook.save(filename)

    if flag:
        flag = False
        return {
            'filename': filename,
            'current_sheet': new_sheet,
            'workbook': workbook,
        }

    self.label_status.setText(f'<font color="green">Файл создан </font>')


def check_current_result_file(self):
    current_datetime = self.ui.date_start.dateTime().toPyDateTime()
    filename = self.path_all['file']  # + current_datetime.strftime("%B %Y.xlsx")

    if os.path.exists(filename):
        workbook = load_workbook(filename)
        sheet_name = f'{self.ui.cb_user.currentText()}' \
                     f'({self.ui.date_start.dateTime().toString("d HH-mm")}-' \
                     f'{self.ui.date_finish.dateTime().toString("d HH-mm")})'

        if not self.ui.check_box_list.isChecked():
            result = {
                'filename': filename,
                'current_sheet': (lambda: workbook[sheet_name] if sheet_name in workbook else None)(),
                'workbook': workbook,
            }
            if result['current_sheet'] is None:
                # print(f'Не найден лист {sheet_name}')
                self.label_status.setText(f'<font color="red">Не найден лист {sheet_name}</font>')
                return 0
            self.path_all['file'] = filename
            self.ui.le_file.setText(filename.split('/')[-1])
            return result

        self.ui.check_box_list.setChecked(False)

        if sheet_name in workbook.sheetnames:
            existing_sheet = workbook[sheet_name]
            workbook.remove(existing_sheet)

        template_sheet = workbook["Шаблон"]
        new_sheet = workbook.copy_worksheet(template_sheet)
        new_sheet.title = sheet_name
        workbook.active = len(workbook.sheetnames) - 1

        template_sheet.sheet_state = "hidden"

        workbook.save(filename)

    else:
        results = create_new_month(self, flag=True)
        new_sheet = results['current_sheet']
        workbook = results['workbook']

    result = {
        'filename': filename,
        'current_sheet': new_sheet,
        'workbook': workbook,
    }

    self.path_all['file'] = filename
    self.ui.le_file.setText(filename.split('/')[-1])

    return result


def find_empty_cell_top(row_index, column_index, sheet):
    # Поиск первой пустой ячейки сверху вниз
    empty_cell = None
    # row_index = 3

    while empty_cell is None:
        cell_value = sheet.cell(row=row_index, column=column_index).value
        if cell_value is None:
            empty_cell = sheet.cell(row=row_index, column=column_index).coordinate
        else:
            row_index += 1

    return coordinate_to_tuple(empty_cell)


def find_empty_cell_bottom(row_index, column_index, sheet):
    # Поиск первой пустой ячейки сверху вниз
    empty_cell = None

    while empty_cell is None:
        cell_value = sheet.cell(row=row_index, column=column_index).value
        if cell_value is None:
            empty_cell = sheet.cell(row=row_index, column=column_index).coordinate
        else:
            row_index -= 1

    return coordinate_to_tuple(empty_cell)


def set_cell(cell, value):
    font = Font(name='Arial', size=12, bold=False, italic=False)
    alignment = Alignment(horizontal='center')
    cell.value = value
    cell.font = font
    cell.alignment = alignment