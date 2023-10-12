from utils.excel.main_file import *


def write_to_excel_exnode_orders(filename, sheet, workbook, filtered_data_exnode_orders):
    """
        Функция write_to_excel_exnode_orders записывает данные о заказах на покупку на бирже Exnode в файл Excel.
    Она принимает следующие аргументы: имя файла, лист, рабочую книгу и отфильтрованные данные о заказах на покупку
    на Exnode.

        Функция ищет пустые ячейки в столбцах, где будут записываться данные. Затем она проходит по каждому элементу
    в отфильтрованных данных и записывает значения в соответствующие ячейки. Для столбца с суммой в USDT используется
    функция find_empty_cell_top для поиска пустой ячейки в столбце 7, начиная с третьей строки. Аналогично, для столбца
    с суммой в RUB используется функция find_empty_cell_top для поиска пустой ячейки в столбце 5, начиная с третьей
    строки.

        Функция также записывает значение 'exnode' в столбец с информацией о рынке, используя функцию
    find_empty_cell_top для поиска пустой ячейки в столбце 4, начиная с третьей строки.

    :param filename:
    :param sheet:
    :param workbook:
    :param filtered_data_exnode_orders:
    :return:
    """
    cell_exnode_buy_usdt = find_empty_cell_top(3, 7, sheet)
    cell_exnode_buy_rub = find_empty_cell_top(3, 5, sheet)
    # cell_exnode_mean = find_empty_cell_top(3, 6, sheet)
    cell_market = find_empty_cell_top(3, 4, sheet)

    for index, (value_usdt, value_rub) in enumerate(
            zip(filtered_data_exnode_orders['usdt'],
                filtered_data_exnode_orders['rub'])):
        cell_usdt = sheet.cell(row=cell_exnode_buy_usdt[0] + index, column=cell_exnode_buy_usdt[1])
        set_cell(cell_usdt, value_usdt)

        cell_rub = sheet.cell(row=cell_exnode_buy_rub[0] + index, column=cell_exnode_buy_rub[1])
        set_cell(cell_rub, value_rub)

        # cell_rub = sheet.cell(row=cell_exnode_mean[0] + index, column=cell_exnode_mean[1])
        # set_cell(cell_rub, value_rub / value_usdt)

        cell_rub = sheet.cell(row=cell_market[0] + index, column=cell_market[1])
        set_cell(cell_rub, 'exnode')

    workbook.save(filename)


def write_to_excel_exnode_komsa(filename, sheet, workbook, filtered_data_exnode_in_out):
    """
        Функция write_to_excel_exnode_komsa записывает данные о комиссии на бирже Exnode в файл Excel. Она принимает
    следующие аргументы: имя файла, лист, рабочую книгу и отфильтрованные данные о комиссии на Exnode.

        Функция вычисляет значение комиссии, умножая количество элементов в отфильтрованных данных на -2. Затем
    она записывает значение комиссии в соответствующую ячейку с помощью функции set_cell.

    :param filename:
    :param sheet:
    :param workbook:
    :param filtered_data_exnode_in_out:
    :return:
    """
    cell_coords_comm = (189, 7)
    value_comm = filtered_data_exnode_in_out.__len__() * -2
    cell_comm = sheet.cell(row=cell_coords_comm[0], column=cell_coords_comm[1])
    set_cell(cell_comm, value_comm)

    workbook.save(filename)
