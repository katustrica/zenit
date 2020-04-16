#!/usr/bin/env Python3
import PySimpleGUIQt as sg
from pandas import DataFrame
from os.path import dirname
from collections import OrderedDict
import pandas as pd
from pathlib import Path
from zeep import Client
from datetime import datetime
from dateutil.parser import parse
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Side

today = datetime.today()
syear, fyear = today.year-1, today.year-1
smonth, fmonth = 2, 2


def get_banks_data_and_name(regnums, table_conf, dates):
    cl = Client("http://cbr.ru/CreditInfoWebServ/CreditOrgInfo.asmx?wsdl")
    f_dates = []
    for date in dates:
        temp = date.split('-')
        year = int(temp[0])
        month = int(temp[1].replace('0', ''))
        if month > 1:
            f_date_month = f'{month - 1:2}'.replace(' ', '0')
            f_date_year = year
        else:
            f_date_month = 12
            f_date_year = year - 1
        f_dates.append(f'{f_date_month}.{f_date_year}')
    row_columns = {table[1]: {} for table in table_conf}
    row_names = []
    for table in table_conf:
        row_columns[table[1]][table[2]] = table[0]
        row_names.append(table[0])

    banks_data = {regnum: {f_date: {} for f_date in f_dates} for regnum in regnums}
    banks_name = {regnum: True for regnum in regnums}
    for regnum in regnums:
        empty = True
        for date, f_date in zip(dates, f_dates):
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
                            banks_data[regnum][f_date][name_row] = element.find(col).text
                if element.tag == 'F1011':
                    banks_name[regnum] = element.find('cname').text
        if empty:
            banks_data[regnum] = empty
            continue
    for name, data in banks_data.items():
        if isinstance(data, bool):
            regnums.remove(name)
    sorted_bank_data = {regnum: {f_date: OrderedDict() for f_date in f_dates} for regnum in regnums}
    for regnum in regnums:
        for f_date in f_dates:
            for row_name in row_names:
                if row_name in banks_data[regnum][f_date].keys():
                    value = banks_data[regnum][f_date][row_name][:-5]
                    sorted_bank_data[regnum][f_date][row_name] = '{0:,}'.format(int(value)).replace(',', ' ')
                else:
                    sorted_bank_data[regnum][f_date][row_name] = 'Нет данных'
    return sorted_bank_data, banks_name


def save_excel_data(banks_data, banks_name):
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

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
                cell.border = thin_border
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column].width = adjusted_width
        worksheet.append(['Название банка', f'{name}'])
        worksheet.move_range(f'A1:{[r for r in worksheet.rows][-1][-1].coordinate}', rows=1)
        last_row_number = [r for r in worksheet.rows][-1][-1].row
        worksheet.move_range(f'A{last_row_number}:B{last_row_number}', rows=-(last_row_number-1))
        wb.save(f'{name}.xlsx')


def to_num(settings):
    setting_num = []
    for setting in settings:
        num = None
        if setting[2] == 'numsc':
            num = 1
        elif setting[2] == 'vr':
            num = 2
        elif setting[2] == 'vv':
            num = 3
        elif setting[2] == 'vitg':
            num = 4
        elif setting[2] == 'ora':
            num = 5
        elif setting[2] == 'ova':
            num = 6
        elif setting[2] == 'oitga':
            num = 7
        elif setting[2] == 'orp':
            num = 8
        elif setting[2] == 'ovp':
            num = 9
        elif setting[2] == 'oitgp':
            num = 10
        elif setting[2] == 'ir':
            num = 11
        elif setting[2] == 'iv':
            num = 12
        elif setting[2] == 'iitg':
            num = 13
        else:
            continue
        setting_num.append([setting[0], setting[1], num])
    return setting_num


def to_str(settings):
    setting_num = []
    for setting in settings:
        st = None
        if setting[2] == 1:
            st = 'numsc'
        elif setting[2] == 2:
            st = 'vr'
        elif setting[2] == 3:
            st = 'vv'
        elif setting[2] == 4:
            st = 'vitg'
        elif setting[2] == 5:
            st = 'ora'
        elif setting[2] == 6:
            st = 'ova'
        elif setting[2] == 7:
            st = 'oitga'
        elif setting[2] == 8:
            st = 'orp'
        elif setting[2] == 9:
            st = 'ovp'
        elif setting[2] == 10:
            st = 'oitgp'
        elif setting[2] == 11:
            st = 'ir'
        elif setting[2] == 12:
            st = 'iv'
        elif setting[2] == 13:
            st = 'iitg'
        else:
            continue
        setting_num.append([setting[0], setting[1], st])
    return setting_num


settings = [
    ['Выдано БГ', '91315', 'oitgp'],
    ['Погашено БГ', '91315', 'oitga'],
    ['Портфель БГ', '91315', 'iitg'],
    ['Комиссия БГ', '47502', 'oitga'],
    ['Выдачи в рамках лимита', '91319', 'oitga'],
    ['Неис лимиты БГ на дату', '91319', 'iitg'],
    ['Выплачено по требованию БГ', '60315', 'oitga'],
]

layout1 = [
    [
        sg.Frame('Номера банков', [[
            sg.Text('№, №'), sg.InputText('1481', key='-REGNUMS-'),
            ]]
        )

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

win1 = sg.Window('101 форма', layout1)
win2_active = False
banks_data, banks_name = {}, {}
dates = []
while True:
    ev1, val1 = win1.read(timeout=100)
    if ev1 in (None, 'Exit'):
        break
    elif ev1 == 'Взять данные':
        while True:
            if (syear < fyear) or (fyear == syear and smonth <= fmonth):
                month = f'{smonth:2}'.replace(' ', '0')
                dates.append(f'{syear}-{month}-01')
                if smonth == 12:
                    smonth=0
                    syear+=1
                else:
                    smonth += 1
            else:
                break

        regnums = list(
            set(val1['-REGNUMS-'].replace(' ', ',').split(sep=','))
        )
        if '' in regnums:
            regnums.remove('')
        # проверить на int str
        banks_data, banks_name = get_banks_data_and_name(regnums, settings, dates)
        regnums_for_listbox = [
            f'{code:>15} | {"нет данных за этот период" if isinstance(name, bool) else name:<20}'.replace(' ', ' ')
            for code, name in banks_name.items()
        ]
        win1['-LISTBOX-'].update(regnums_for_listbox)
    elif ev1 == 'Сохранить':
        if banks_data and banks_name:
            save_excel_data(banks_data, banks_name)
    elif ev1 == 'Удалить':
        if val1['-LISTBOX-']:
            for key in [value.split()[0] for value in val1['-LISTBOX-']]:
                try:
                    del banks_data[key]
                    del banks_name[key]
                except KeyError:
                    continue
        regnums_for_listbox = [
            f'{code:>8} | {"нет данных за этот период" if isinstance(name, bool) else name:<20}'.replace(' ', ' ')
            for code, name in banks_name.items()
        ]
        win1['-LISTBOX-'].update(regnums_for_listbox)
    elif ev1 == 'Настройки парсинга':
        win2_active = True
        win1.Hide()
        layout2 = [
            [
                sg.Frame('Текущие настройки', [
                        [
                            sg.Listbox(values=[
                                    f'{setting[0]:<40} | {setting[1]:>6} | {setting[2]:<3}' for setting in to_num(settings)
                                ], select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, key='-SLISTBOX-')
                        ],
                        [sg.Button('Вверх'), sg.Button('Вниз'), sg.Button('Удалить')]
                    ]
                )
            ],
            [
                sg.Frame('(Имя столбца, номер строки, номер столбца)', [
                        [sg.Text('Номера банков'), sg.InputText('', key='-SETTINGS-')],
                        [sg.Button('Добавить')]
                    ]
                )
            ],
            [
                sg.Frame('Начало', [[
                        sg.Text('Год'), sg.InputText(f'{syear}', size=(45, 20), key='-SYEAR-'),
                        sg.Text('Месяц'), sg.InputText(f'{smonth-1}', size=(30, 20), key='-SMONTH-'),
                    ]],
                ),
                sg.Frame('Конец', [[
                        sg.Text('Год'), sg.InputText(f'{fyear}', size=(45, 20), key='-FYEAR-'),
                        sg.Text('Месяц'), sg.InputText(f'{fmonth-1}', size=(30, 20), key='-FMONTH-')
                    ]]
                )
            ],
            [sg.Button('Сохранить даты')]
        ]
        win2 = sg.Window('Настройки', layout2)
        while True:
            ev2, val2 = win2.Read()
            if ev2 in (None, 'Exit'):
                win2.close()
                win2_active = False
                win1.UnHide()
                break
            if ev2 == 'Добавить':
                # filtered settings
                f_s = []
                # readed settings
                r_s = re.split(r'[()]', val2['-SETTINGS-'])
                for elem in r_s:
                    element = elem.split(', ')
                    if len(element) == 3:
                        try:
                            f_s.append([element[0], element[1], int(element[2])])
                        except (ValueError, IndexError):
                            continue
                settings += to_str(f_s)
                lis = [f'{sett[0]:<20} | {sett[1]:>6} | {sett[2]:<3}' for sett in to_num(settings)]
                win2['-SLISTBOX-'].update(lis)
            if ev2 == 'Вверх':
                if val2['-SLISTBOX-']:
                    if len(val2['-SLISTBOX-']) > 1:
                        continue
                    key = val2['-SLISTBOX-'][0].split('|')[0]
                    zero_set = " ".join([el for el in key.strip().split(' ') if el.strip()])
                    i = 0
                    for setting in settings:
                        if setting[0] == zero_set:
                            if i == 0:
                                continue
                            settings[i], settings[i-1] = settings[i-1], settings[i]
                            break
                        i += 1
                lis = [f'{sett[0]:<20} | {sett[1]:>6} | {sett[2]:<3}' for sett in to_num(settings)]
                win2['-SLISTBOX-'].update(lis)
            if ev2 == 'Вниз':
                if val2['-SLISTBOX-']:
                    if len(val2['-SLISTBOX-']) > 1:
                        continue
                    key = val2['-SLISTBOX-'][0].split('|')[0]
                    zero_set = " ".join([el for el in key.strip().split(' ') if el.strip()])
                    i = 0
                    for setting in settings:
                        if setting[0] == zero_set:
                            if i == len(settings)-1:
                                continue
                            settings[i], settings[i+1] = settings[i+1], settings[i]
                            break
                        i += 1
                lis = [f'{sett[0]:<20} | {sett[1]:>6} | {sett[2]:<3}' for sett in to_num(settings)]
                win2['-SLISTBOX-'].update(lis)
            if ev2 == 'Удалить':
                if val2['-SLISTBOX-']:
                    for key in [value.split('|')[0] for value in val2['-SLISTBOX-']]:
                        zero_set = " ".join([el for el in key.strip().split(' ') if el.strip()])
                        for setting in settings:
                            if setting[0] == zero_set:
                                settings.remove(setting)
                                break
                lis = [f'{sett[0]:<20} | {sett[1]:>6} | {sett[2]:<3}' for sett in to_num(settings)]
                win2['-SLISTBOX-'].update(lis)
            if ev2 == 'Сохранить даты':
                syear, fyear = int(val2['-SYEAR-']), int(val2['-FYEAR-'])
                smonth, fmonth = int(val2['-SMONTH-']), int(val2['-FMONTH-'])
                if smonth == 12:
                    smonth = 0
                    syear += 1
                else:
                    smonth += 1
                if fmonth == 12:
                    fmonth = 0
                    fyear += 1
                else:
                    fmonth += 1
win1.close()