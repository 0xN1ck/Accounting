from utils.excel.main_file import *


def write_to_excel_exnode_history(filename, sheet, workbook, filtered_data_exnode_history):
    cell_exnode_buy_usdt = find_empty_cell_top(3, 7, sheet)
    cell_exnode_buy_rub = find_empty_cell_top(3, 5, sheet)
    cell_exnode_mean = find_empty_cell_top(3, 6, sheet)
    cell_market = find_empty_cell_top(3, 4, sheet)

    for index, (value_usdt, value_rub) in enumerate(
            zip(filtered_data_exnode_history.loc[filtered_data_exnode_history['direction'] == 'IN', 'usdt'],
                filtered_data_exnode_history.loc[filtered_data_exnode_history['direction'] == 'IN', 'rub'])):
        cell_usdt = sheet.cell(row=cell_exnode_buy_usdt[0] + index, column=cell_exnode_buy_usdt[1])
        set_cell(cell_usdt, value_usdt)

        cell_rub = sheet.cell(row=cell_exnode_buy_rub[0] + index, column=cell_exnode_buy_rub[1])
        set_cell(cell_rub, value_rub)

        cell_rub = sheet.cell(row=cell_exnode_mean[0] + index, column=cell_exnode_mean[1])
        set_cell(cell_rub, value_rub / value_usdt)

        cell_rub = sheet.cell(row=cell_market[0] + index, column=cell_market[1])
        set_cell(cell_rub, 'EXNODE')

    workbook.save(filename)


def write_to_excel_exnode_komsa(filename, sheet, workbook, filtered_data_exnode_history):
    cell_coords_comm = (189, 7)
    value_comm = filtered_data_exnode_history.loc[filtered_data_exnode_history['direction'] == 'OUT', 'commission'].sum() * -1
    cell_comm = sheet.cell(row=cell_coords_comm[0], column=cell_coords_comm[1])
    set_cell(cell_comm, value_comm)

    workbook.save(filename)
