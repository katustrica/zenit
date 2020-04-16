from zeep import Client
from datetime import datetime, timedelta
from dateutil.parser import parse
import pandas as pd
from pprint import pprint
from lxml import etree


regnums = [1481, 415222]
table_conf = [
    ['Выдано БГ', '91315', 'oitgp'],
    ['Погашено БГ', '91315', 'oitga'],
    ['Портфель БГ', '91315', 'iitg'],
    ['Комиссия БГ', '47502', 'oitga'],
    ['Выдачи в рамках лимита', '91319', 'oitga'],
    ['Неис лимиты БГ на дату', '91319', 'iitg'],
    ['Выплачено по требованию БГ', '60315', 'oitga'],
]


def get_banks_data_and_name(regnums, table_conf, dates):
    row_columns = {table[1]: {} for table in table_conf}
    row_names = []
    for table in table_conf:
        row_columns[table[1]][table[2]] = table[0]
        row_names.append(table[0])

    banks_data = {regnum: {date: {} for date in dates} for regnum in regnums}
    banks_name = {regnum: True for regnum in regnums}
    for regnum in regnums:
        empty = True
        for date in dates:
            dt = parse(date)
            try:
                data = cl.service.Data101FNewXML(CredorgNumber=regnum, Dt=dt)
            except:
                banks_data[regnum] = empty
                continue
            for element in data:
                if element.tag == 'F101':
                    empty = False
                    numsc = element.find('numsc').text
                    if numsc in row_columns.keys():
                        column_names = row_columns[numsc].keys()
                        for col in column_names:
                            name_row = row_columns[numsc][col]
                            banks_data[regnum][date][name_row] = element.find(col).text
                if element.tag == 'F1011':
                    banks_name[regnum] = element.find('cname').text
        if empty:
            banks_data[regnum] = empty
            continue
    for name, data in banks_data.items():
        if isinstance(data, bool):
            regnums.remove(name)
    sorted_bank_data = {regnum: {date: OrderedDict() for date in dates} for regnum in regnums}
    for regnum in regnums:
        for date in dates:
            for row_name in row_names:
                if row_name in banks_data[regnum][date].keys():
                    value = banks_data[regnum][date][row_name][:-5]
                    sorted_bank_data[regnum][date][row_name] = '{0:,}'.format(int(value)).replace(',', ' ')
                else:
                    sorted_bank_data[regnum][date][row_name] = 'Нет данных'
    return sorted_bank_data, banks_name


def save_excel_data(banks_data, banks_name):
    for regnum, data in banks_data.items():
        name = banks_name[regnum]
        if name is True:
            continue
        df = pd.DataFrame.from_dict(data)
        name = name.replace('"', '')
        df.to_excel(f'{name}.xlsx')
        wb = openpyxl.load_workbook(f'{name}.xlsx')
        worksheet = wb.active
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.column_letter == 'A':
                    cell.alignment = Alignment(horizontal='left')
                elif cell.row != 1:
                    cell.alignment = Alignment(horizontal='right')

                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column].width = adjusted_width

        wb.save(f'{name}.xlsx')

banks_data, banks_name = get_banks_data_and_name(regnums, table_conf, dates)
save_excel_data(banks_data, banks_name)