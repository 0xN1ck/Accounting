import pandas as pd
from utils.excel.main_file import *


def write_to_excel_bybit_history_p2p(filename, sheet, workbook, filtered_data_bybit_history_p2p_buy):
    """
            Функция write_to_excel_bybit_history_p2p(filename, sheet, workbook, filtered_data_bybit_history_p2p_buy)
        записывает данные о покупках криптовалюты на бирже Bybit в файл Excel. Функция принимает следующие параметры:

        - filename: имя файла Excel, в который будут записаны данные,
        - sheet: лист Excel, на котором будут записаны данные,
        - workbook: рабочая книга Excel,
        - filtered_data_bybit_history_p2p_buy: отфильтрованные данные о покупках криптовалюты на бирже Bybit.

            Функция ищет пустые ячейки на листе Excel и записывает данные о покупках криптовалюты на бирже Bybit
        в соответствующие ячейки. Данные записываются по разделенным по монетам координатам. Для каждой монеты
        записываются следующие значения: общая стоимость в рублях, цена за единицу криптовалюты и количество
        криптовалюты. Затем функция сохраняет изменения в файл Excel.

    :param filename:
    :param sheet:
    :param workbook:
    :param filtered_data_bybit_history_p2p_buy:
    :return:
    """
    cell_bybit_buy_usdt = find_empty_cell_top(3, 7, sheet)
    # cell_bybit_buy_mean = find_empty_cell_top(3, 6, sheet)
    cell_bybit_buy_rub = find_empty_cell_top(3, 5, sheet)
    cell_bybit_buy_market = find_empty_cell_top(3, 4, sheet)

    index = {
        'BTC': 0,
        'ETH': 0,
        'USDT': 0
    }
    start_coords_coins = {
        'BTC': {
            'common_price': (210, 1),
            'price': (210, 2),
            'amount': (210, 3)
        },
        'ETH': {
            'common_price': (253, 1),
            'price': (253, 2),
            'amount': (253, 3)
        }
    }
    one_coin = False

    for name_coin in filtered_data_bybit_history_p2p_buy.index:
        if type(filtered_data_bybit_history_p2p_buy.loc[name_coin]['Fiat Amount']) != pd.Series:
            one_coin = True
        if not name_coin == 'USDT':
            value_common_price = filtered_data_bybit_history_p2p_buy.loc[name_coin]['Fiat Amount'].iloc[index[name_coin]] \
                if not one_coin else filtered_data_bybit_history_p2p_buy.loc[name_coin]['Fiat Amount']
            cell_coords_common_price = find_empty_cell_top(start_coords_coins[name_coin]['common_price'][0],
                                                           start_coords_coins[name_coin]['common_price'][1], sheet)
            cell_common_price = sheet.cell(row=cell_coords_common_price[0], column=cell_coords_common_price[1])
            set_cell(cell_common_price, value_common_price)

            value_price = filtered_data_bybit_history_p2p_buy.loc[name_coin]['Price'].iloc[index[name_coin]] \
                if not one_coin else filtered_data_bybit_history_p2p_buy.loc[name_coin]['Price']
            cell_coords_price = find_empty_cell_top(start_coords_coins[name_coin]['price'][0],
                                                    start_coords_coins[name_coin]['price'][1], sheet)
            cell_price = sheet.cell(row=cell_coords_price[0], column=cell_coords_price[1])
            set_cell(cell_price, value_price)

            value_amount = filtered_data_bybit_history_p2p_buy.loc[name_coin]['Coin Amount'].iloc[index[name_coin]] \
                if not one_coin else filtered_data_bybit_history_p2p_buy.loc[name_coin]['Coin Amount']
            cell_coords_amount = find_empty_cell_top(start_coords_coins[name_coin]['amount'][0],
                                                     start_coords_coins[name_coin]['amount'][1], sheet)
            cell_amount = sheet.cell(row=cell_coords_amount[0], column=cell_coords_amount[1])
            set_cell(cell_amount, value_amount)

            workbook.save(filename)
            index[name_coin] += 1
            one_coin = False
            continue

        value_rub = filtered_data_bybit_history_p2p_buy.loc[name_coin]['Fiat Amount'].iloc[index['USDT']] \
            if not one_coin else filtered_data_bybit_history_p2p_buy.loc[name_coin]['Fiat Amount']

        value_usdt = filtered_data_bybit_history_p2p_buy.loc[name_coin]['Coin Amount'].iloc[index['USDT']] \
            if not one_coin else filtered_data_bybit_history_p2p_buy.loc[name_coin]['Coin Amount']

        # value_mean = filtered_data_bybit_history_p2p_buy.loc[name_coin]['Price'].iloc[index['USDT']] \
        #     if not one_coin else filtered_data_bybit_history_p2p_buy.loc[name_coin]['Price']

        cell_usdt = sheet.cell(row=cell_bybit_buy_usdt[0] + index['USDT'], column=cell_bybit_buy_usdt[1])
        set_cell(cell_usdt, value_usdt)

        cell_rub = sheet.cell(row=cell_bybit_buy_rub[0] + index['USDT'], column=cell_bybit_buy_rub[1])
        set_cell(cell_rub, value_rub)

        # cell_mean = sheet.cell(row=cell_bybit_buy_mean[0] + index['USDT'], column=cell_bybit_buy_mean[1])
        # set_cell(cell_mean, value_mean)

        cell_market = sheet.cell(row=cell_bybit_buy_market[0] + index['USDT'], column=cell_bybit_buy_market[1])
        set_cell(cell_market, 'bybit')

        value_rub, value_usdt = 0, 0
        index['USDT'] += 1
        one_coin = False

    workbook.save(filename)


def write_to_excel_bybit_history_p2p_sell(filename, sheet, workbook, filtered_data_bybit_history_p2p_sell):
    """
            Функция write_to_excel_bybit_history_p2p_sell(filename, sheet, workbook, filtered_data_bybit_history_p2p_sell)
        записывает данные о продажах криптовалюты на бирже Bybit в файл Excel. Функция принимает следующие параметры:

            - filename: имя файла Excel, в который будут записаны данные,
            - sheet: лист Excel, на котором будут записаны данные,
            - workbook: рабочая книга Excel,
            - filtered_data_bybit_history_p2p_sell: отфильтрованные данные о продажах криптовалюты на бирже Bybit.

            Функция ищет пустые ячейки на листе Excel и записывает данные о продажах криптовалюты на бирже Bybit
        в соответствующие ячейки. Данные записываются по разделенным по монетам координатам. Для каждой монеты
        записываются следующие значения: количество проданной криптовалюты, общая стоимость в рублях и средняя цена
        за единицу криптовалюты. Затем функция сохраняет изменения в файл Excel.

    :param filename:
    :param sheet:
    :param workbook:
    :param filtered_data_bybit_history_p2p_sell:
    :return:
    """
    cell_bybit_sell_usdt = find_empty_cell_top(3, 12, sheet)
    # cell_bybit_sell_mean = find_empty_cell_top(3, 11, sheet)
    cell_bybit_sell_rub = find_empty_cell_top(3, 10, sheet)
    cell_bybit_sell_market = find_empty_cell_top(3, 9, sheet)

    for index, value in enumerate(filtered_data_bybit_history_p2p_sell.index):
        cell_usdt = sheet.cell(row=cell_bybit_sell_usdt[0] + index, column=cell_bybit_sell_usdt[1])
        set_cell(cell_usdt, filtered_data_bybit_history_p2p_sell.loc[value]['Coin Amount'])

        cell_rub = sheet.cell(row=cell_bybit_sell_rub[0] + index, column=cell_bybit_sell_rub[1])
        set_cell(cell_rub, filtered_data_bybit_history_p2p_sell.loc[value]['Fiat Amount'])

        # cell_mean = sheet.cell(row=cell_bybit_sell_mean[0] + index, column=cell_bybit_sell_mean[1])
        # set_cell(cell_mean, filtered_data_bybit_history_p2p_sell.loc[value]['Price'])

        cell_market = sheet.cell(row=cell_bybit_sell_market[0] + index, column=cell_bybit_sell_market[1])
        set_cell(cell_market, 'Bybit')

    workbook.save(filename)