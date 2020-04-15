#!/usr/bin/env Python3
import PySimpleGUIQt as sg
from pandas import read_html, DataFrame
from os.path import dirname
from bs4 import BeautifulSoup
from requests import Session
from pathlib import Path
from zeep import Client
from datetime import datetime, timedelta
from dateutil.parser import parse
from lxml import etree


# establishing session
today = datetime.today()

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


# sg.change_look_and_feel('DefaultNoMoreNagging')
column1 = [[sg.Listbox(
    values=[],
    select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE,
    key='-LISTBOX-'
    )
]]
layout1 = [
    [
        sg.Text('Номера банков'), sg.InputText('415, 1481 1326', key='-REGNUMS-'),
        sg.Text('Год',), sg.InputText(f'{datetime.today():%Y}', size=(45, 20), key='-YEAR-'),
        sg.Text('Месяц'), sg.InputText(f'{datetime.today():%m}', size=(30, 20), key='-MONTH-')
    ],
    [sg.Button('Настройки парсинга')],
    [
        sg.Frame('Список банков', [[
            sg.Listbox(values=[], select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, key='-LISTBOX-')
            ]]
        )
    ],
    [sg.Button('Взять данные'), sg.Text(' ' * 79), sg.Button('Удалить')],
    [
        sg.Frame('Cохранение', [[
                sg.FolderBrowse('Выберете папку', target='-PATH-'),
                sg.InputText(f'{dirname(__file__)}', key='-PATH-'),
                sg.Button('Сохранить')
            ]]
        )
    ],
    [sg.ProgressBar(1000, orientation='h', key='progressbar')]
]
layout2 = [
    [sg.Text('Номера банков'), sg.InputText('(Выдано БГ, 91315, 1)', key='-REGNUMS-')],

    [sg.Button('Сохранить')]
]

win1 = sg.Window('101 форма', layout1)
win2_active = False

while True:
    event, values = win1.read(timeout=100)
    year = values['-YEAR-']
    month = values['-MONTH-']
    if event in (None, 'Exit'):
        break
    elif event == 'Взять данные':
        regnums = list(
            set(values['-REGNUMS-'].replace(' ', ',').split(sep=','))
        )
        if '' in regnums:
            regnums.remove('')
        # проверить на int str
        banks_data, banks_name = get_banks_data_and_name(regnums, table_conf, dates)
        regnums_for_listbox = [
            f'{code:>15} | {name:<20} {"| нет данных за этот период" if isinstance(tables_df[code], bool) else " ":<30}'.replace(' ', ' ').lower()
            for code, name in banks_name_and_code.items()
        ]
    win1['-LISTBOX-'].update(regnums_for_listbox)
    elif event == 'Сохранить':
        save_excel_data(banks_data, banks_name)
    elif event == 'Удалить':
        if values['-LISTBOX-']:
            for key in [value.split()[0] for value in values['-LISTBOX-']]:
                del banks_data[key]
                del banks_name[key]
    elif event == 'Launch 2' and not win2_active:
        win2_active = True
        win1.Hide()
        win2 = sg.Window('Window 2', layout2)
        while True:
            ev2, vals2 = win2.Read()
            if event in (None, 'Exit', 'Сохранить'):
                win2.close()
                win2_active = False
                win1.UnHide()
                break
win1.close()