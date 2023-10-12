from utils.excel.main_file import *
import pandas as pd


def write_to_to_excel_huobi_history_p2p_buy(filename, sheet, workbook, huobi_history_p2p_buy_sum):
    """
            Функция write_to_to_excel_huobi_history_p2p_buy используется для записи данных о покупках на бирже Huobi
        в файл Excel. Она принимает следующие аргументы:
        - filename - имя файла Excel, в который нужно записать данные
        - sheet - лист Excel, на котором нужно записать данные
        - workbook - рабочая книга Excel, в которую нужно записать данные
        - huobi_history_p2p_buy_sum - DataFrame с данными о покупках на бирже Huobi

            Функция ищет пустые ячейки на листе Excel и записывает данные о покупках в соответствующие ячейки.
        Данные о покупках представлены в DataFrame huobi_history_p2p_buy_sum, который содержит информацию о цене,
        количестве и общей сумме покупок для каждой монеты. Для каждой монеты данные записываются в отдельные
        ячейки, а для монеты USDT данные записываются в ячейки, которые относятся к покупкам USDT.

    :param filename:
    :param sheet:
    :param workbook:
    :param huobi_history_p2p_buy_sum:
    :return:
    """
    cell_huobi_buy_usdt = find_empty_cell_top(3, 7, sheet)
    # cell_huobi_buy_mean = find_empty_cell_top(3, 6, sheet)
    cell_huobi_buy_rub = find_empty_cell_top(3, 5, sheet)
    cell_huobi_buy_market = find_empty_cell_top(3, 4, sheet)

    huobi_history_p2p_buy_sum.set_index('Монета', inplace=True)
    index = {
        'BTC': 0,
        'ETH': 0,
        'USDT': 0
    }
    start_coords_coins = {
        'BTC': {
            'common_price': (210, 7),
            'price': (210, 8),
            'amount': (210, 9)
        },
        'ETH': {
            'common_price': (253, 7),
            'price': (253, 8),
            'amount': (253, 9)
        }
    }
    one_coin = False

    for name_coin in huobi_history_p2p_buy_sum.index:
        if type(huobi_history_p2p_buy_sum.loc[name_coin]['Общая цена']) != pd.Series:
            one_coin = True

        if not name_coin == 'USDT':
            value_common_price = huobi_history_p2p_buy_sum.loc[name_coin]['Общая цена'].iloc[index[name_coin]] \
                if not one_coin else huobi_history_p2p_buy_sum.loc[name_coin]['Общая цена']
            cell_coords_common_price = find_empty_cell_top(start_coords_coins[name_coin]['common_price'][0],
                                                           start_coords_coins[name_coin]['common_price'][1], sheet)
            cell_common_price = sheet.cell(row=cell_coords_common_price[0], column=cell_coords_common_price[1])
            set_cell(cell_common_price, value_common_price)

            value_price = huobi_history_p2p_buy_sum.loc[name_coin]['Цена за ед.'].iloc[index[name_coin]] \
                if not one_coin else huobi_history_p2p_buy_sum.loc[name_coin]['Цена за ед.']
            cell_coords_price = find_empty_cell_top(start_coords_coins[name_coin]['price'][0],
                                                    start_coords_coins[name_coin]['price'][1], sheet)
            cell_price = sheet.cell(row=cell_coords_price[0], column=cell_coords_price[1])
            set_cell(cell_price, value_price)

            value_amount = huobi_history_p2p_buy_sum.loc[name_coin]['Количество'].iloc[index[name_coin]] \
                if not one_coin else huobi_history_p2p_buy_sum.loc[name_coin]['Количество']
            cell_coords_amount = find_empty_cell_top(start_coords_coins[name_coin]['amount'][0],
                                                     start_coords_coins[name_coin]['amount'][1], sheet)
            cell_amount = sheet.cell(row=cell_coords_amount[0], column=cell_coords_amount[1])
            set_cell(cell_amount, value_amount)

            workbook.save(filename)
            index[name_coin] += 1
            one_coin = False
            continue

        value_rub = huobi_history_p2p_buy_sum.loc[name_coin]['Общая цена'].iloc[index['USDT']] \
            if not one_coin else huobi_history_p2p_buy_sum.loc[name_coin]['Общая цена']

        value_usdt = huobi_history_p2p_buy_sum.loc[name_coin]['Количество'].iloc[index['USDT']] \
            if not one_coin else huobi_history_p2p_buy_sum.loc[name_coin]['Количество']

        # value_mean = value_rub / value_usdt

        cell_usdt = sheet.cell(row=cell_huobi_buy_usdt[0] + index['USDT'], column=cell_huobi_buy_usdt[1])
        set_cell(cell_usdt, value_usdt)

        cell_rub = sheet.cell(row=cell_huobi_buy_rub[0] + index['USDT'], column=cell_huobi_buy_rub[1])
        set_cell(cell_rub, value_rub)

        # cell_mean = sheet.cell(row=cell_huobi_buy_mean[0] + index['USDT'], column=cell_huobi_buy_mean[1])
        # set_cell(cell_mean, value_mean)

        cell_market = sheet.cell(row=cell_huobi_buy_market[0] + index['USDT'], column=cell_huobi_buy_market[1])
        set_cell(cell_market, 'huobi')

        value_rub, value_usdt = 0, 0
        index['USDT'] += 1
        one_coin = False

    workbook.save(filename)


def write_to_to_excel_huobi_history_p2p_sell(filename, sheet, workbook, filtered_data_huobi_history_p2p_sell):
    """
            Функция  write_to_to_excel_huobi_history_p2p_sell  используется для записи данных о продажах на бирже Huobi
        в файл Excel. Она принимает следующие аргументы:
        -  filename  - имя файла Excel, в который нужно записать данные
        -  sheet  - лист Excel, на котором нужно записать данные
        -  workbook  - рабочая книга Excel, в которую нужно записать данные
        -  filtered_data_huobi_history_p2p_sell  - DataFrame с данными о продажах на бирже Huobi

            Функция ищет пустые ячейки на листе Excel и записывает данные о продажах в соответствующие ячейки.
        Данные о продажах представлены в DataFrame  filtered_data_huobi_history_p2p_sell, который содержит информацию
        о цене, количестве и общей сумме продаж для каждой монеты. Для каждой монеты данные записываются в отдельные
        ячейки.

    :param filename:
    :param sheet:
    :param workbook:
    :param filtered_data_huobi_history_p2p_sell:
    :return:
    """
    cell_huobi_sell_usdt = find_empty_cell_top(3, 12, sheet)
    # cell_huobi_sell_mean = find_empty_cell_top(3, 11, sheet)
    cell_huobi_sell_rub = find_empty_cell_top(3, 10, sheet)
    cell_huobi_sell_market = find_empty_cell_top(3, 9, sheet)

    for index, value in enumerate(filtered_data_huobi_history_p2p_sell.index):
        cell_usdt = sheet.cell(row=cell_huobi_sell_usdt[0] + index, column=cell_huobi_sell_usdt[1])
        set_cell(cell_usdt, filtered_data_huobi_history_p2p_sell.loc[value]['Количество'])

        cell_rub = sheet.cell(row=cell_huobi_sell_rub[0] + index, column=cell_huobi_sell_rub[1])
        set_cell(cell_rub, filtered_data_huobi_history_p2p_sell.loc[value]['Общая цена'])

        # cell_mean = sheet.cell(row=cell_huobi_sell_mean[0] + index, column=cell_huobi_sell_mean[1])
        # set_cell(cell_mean, filtered_data_huobi_history_p2p_sell.loc[value]['Цена за ед.'])

        cell_market = sheet.cell(row=cell_huobi_sell_market[0] + index, column=cell_huobi_sell_market[1])
        set_cell(cell_market, 'huobi')

    workbook.save(filename)
