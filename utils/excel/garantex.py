from utils.excel.main_file import *


def write_to_excel_gara_history_p2p(filename, sheet, workbook, filtered_data_gara_history_p2p):
    cell_otc_gar = find_empty_cell_top(3, 2, sheet)
    cell_turnover = find_empty_cell_top(3, 1, sheet)

    for index, (value_res, value_total) in enumerate(
            zip(filtered_data_gara_history_p2p['Результат'], filtered_data_gara_history_p2p['Сумма'])):
        cell_res = sheet.cell(row=cell_otc_gar[0] + index, column=cell_otc_gar[1])
        set_cell(cell_res, value_res)

        cell_total = sheet.cell(row=cell_turnover[0] + index, column=cell_turnover[1])
        set_cell(cell_total, value_total)

    workbook.save(filename)


def write_to_excel_gara_swap_history(filename, sheet, workbook, filtered_data_gara_history_swap):
    cell_buy_gar_rub = find_empty_cell_top(3, 5, sheet)
    cell_buy_gar_usdt = find_empty_cell_top(3, 7, sheet)
    cell_buy_gar_course = find_empty_cell_top(3, 6, sheet)
    cell_buy_gar_market = find_empty_cell_top(3, 4, sheet)

    # Запись данных из DataFrame в Excel
    for index, (value_gave, value_get, value_course) in enumerate(zip(filtered_data_gara_history_swap['Отдал'],
                                                                      filtered_data_gara_history_swap['Получил'],
                                                                      filtered_data_gara_history_swap['Средняя цена'])):
        cell_gave = sheet.cell(row=cell_buy_gar_rub[0] + index, column=cell_buy_gar_rub[1])
        set_cell(cell_gave, value_gave)

        cell_get = sheet.cell(row=cell_buy_gar_usdt[0] + index, column=cell_buy_gar_usdt[1])
        set_cell(cell_get, value_get)

        cell_course = sheet.cell(row=cell_buy_gar_course[0] + index, column=cell_buy_gar_course[1])
        set_cell(cell_course, value_course)

        cell_market = sheet.cell(row=cell_buy_gar_market[0] + index, column=cell_buy_gar_market[1])
        set_cell(cell_market, 'garantex')

    workbook.save(filename)


def write_to_excel_gara_account_history(filename, sheet, workbook, filtered_data_gara_history_account):
    cell_coords_comm = (190, 7)
    cell_comm = sheet.cell(row=cell_coords_comm[0], column=cell_coords_comm[1])
    value_cell_comm = filtered_data_gara_history_account['Комиссия'].sum() * -1
    set_cell(cell_comm, value_cell_comm)

    workbook.save(filename)
