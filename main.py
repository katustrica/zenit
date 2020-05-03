import re
from pathlib import Path
from os.path import dirname
from datetime import datetime
from dateutil.parser import parse
from collections import OrderedDict

import pandas as pd
import numpy as np
from zeep import Client
import PySimpleGUIQt as sg

import openpyxl
from openpyxl.styles import Alignment, Border, Side

today = datetime.today()
syear, fyear = today.year-1, today.year-1
smonth, fmonth = 1, 1

temp_syear, temp_fyear = today.year-1, today.year-1
temp_smonth, temp_fmonth = 2, 2
months_name = {
    '01': 'Январь',
    '02': 'Февраль',
    '03': 'Март',
    '04': 'Апрель',
    '05': 'Май',
    '06': 'Июнь',
    '07': 'Июль',
    '08': 'Август',
    '09': 'Сентябрь',
    '10': 'Октябрь',
    '11': 'Ноябрь',
    '12': 'Декабрь'
}
color_popup = '#b1b6fa'
color_popup_ok = '#c5ffc2'
color_popup_info = '#fcfcd2'
sg.ChangeLookAndFeel('TanBlue')
highlight_stroke = 'Выдано БГ'

def get_banks_data_and_name(regnums, table_conf, dates):
    cl = Client("http://cbr.ru/CreditInfoWebServ/CreditOrgInfo.asmx?wsdl")
    f_dates = []
    for date in dates:
        temp = date.split('-')
        year = int(temp[0])
        month = int(temp[1])

        if month > 1:
            f_date_month = f'{month - 1:2}'.replace(' ', '0')
            f_date_year = year
        else:
            f_date_month = '12'
            f_date_year = year - 1
        f_dates.append(f'{months_name[f_date_month]} {f_date_year}')
    row_columns = {table[1]: {} for table in table_conf}
    row_columns_divide = {table[1]: {} for table in table_conf}
    row_names = []

    for table in table_conf:
        if len(table) == 3:
            row_columns[table[1]][table[2]] = table[0]
            row_names.append(table[0])
        else:
            row_names.append(table[0])
            row_columns_divide[table[1]][table[2]] = table[0]
            row_columns_divide[table[3]][table[4]] = table[0]
    row_columns_divide = {key: value for key, value in row_columns_divide.items() if value}
    row_columns = {key: value for key, value in row_columns.items() if value}

    banks_data = {regnum: {f_date: {} for f_date in f_dates} for regnum in regnums}
    banks_name = {regnum: True for regnum in regnums}
    i = 0
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
                    if numsc in row_columns_divide.keys():
                        column_names = row_columns_divide[numsc].keys()
                        for col in column_names:
                            name_row = row_columns_divide[numsc][col]
                            if name_row in banks_data[regnum][f_date].keys():
                                temp_data = float(banks_data[regnum][f_date][name_row])
                                try:
                                    banks_data[regnum][f_date][name_row] = round(temp_data / float(element.find(col).text) * 100, 2)
                                except (TypeError, ZeroDivisionError):
                                    pass
                            else:
                                banks_data[regnum][f_date][name_row] = element.find(col).text
                if element.tag == 'F1011':
                    banks_name[regnum] = element.find('cname').text
        i+=1
        sg.popup_quick_message(f'{i} из {len(regnums)}. Взял данные - №{regnum}', background_color=color_popup, auto_close_duration=2, no_titlebar=True)
        if empty:
            banks_data[regnum] = empty
            continue
    for name, data in banks_data.items():
        if isinstance(data, bool):
            regnums.remove(name)
    sorted_bank_data = {regnum: {f_date: OrderedDict() for f_date in f_dates} for regnum in regnums}
    row_columns_names = sum([list(row.values()) for row in row_columns.values()], [])
    for regnum in regnums:
        for f_date in f_dates:
            for row_name in row_names:
                if row_name in banks_data[regnum][f_date].keys():
                    value = banks_data[regnum][f_date][row_name]
                    if isinstance(value, str) and row_name in row_columns_names:
                        value = value[:-5]
                        sorted_bank_data[regnum][f_date][row_name] = '{0:,}'.format(int(value)).replace(',', ' ')
                    elif isinstance(value, float):
                        sorted_bank_data[regnum][f_date][row_name] = f'{value}%'
                    else:
                        sorted_bank_data[regnum][f_date][row_name] = '-'
                else:
                    sorted_bank_data[regnum][f_date][row_name] = '-'

    return sorted_bank_data, banks_name

def save_excel_data(banks_data, banks_name, path):

    date = datetime.now().strftime("%H %M %S %d-%m-%Y")
    writer = pd.ExcelWriter(path / f'Результаты {date}.xlsx', engine='xlsxwriter')

    banks_data_df = {}
    for regnum, data in banks_data.items():
        name = banks_name[regnum]
        if name is True:
            continue
        df = pd.DataFrame.from_dict(data).reindex([setting[0] for setting in settings])
        for setting in growth_settings:
            if len(df.iloc[setting[1]-1]) < 2:
                continue
            try:
                first_growth = np.array([float(value.replace(' ', '')) for value in df.iloc[setting[1]-1]][:-1])
                second_growth = np.array([float(value.replace(' ', '')) for value in df.iloc[setting[1]-1]][1:])
                growht = (second_growth / first_growth) - 1
                df.loc[setting[0]] = ['']+[f'{num:.2f}%' for num in np.around(growht, 3)*100]
            except Exception as e:
                sg.popup_quick_message(f'Не корректные данные для составления роста у банка - {name}', background_color=color_popup_info, no_titlebar=True)
        banks_data_df[name] = df
    df_s_list = []
    for i, (name, df) in enumerate(banks_data_df.items()):
        df_f = df.replace(r'\s+', '',regex=True)
        df_f = df_f.applymap(lambda x: int(x) if ('%' not in x and x not in ['-', '-', '']) else x)
        df_f.to_excel(writer, sheet_name='Sheet1', startrow=i*(len(settings) + len(growth_settings) + 3) + 1, startcol=1, index=False)
        df_s = pd.DataFrame.from_dict({a: [
                int(b.replace(' ', '')) if (('%' not in b) and (b not in ['-', '']) and isinstance(b, str)) else b
            ] for a, b in df.loc[highlight_stroke].items()}
        )
        df_s.index = [name]
        df_s['Итого'] = sum([num for num in df_s.iloc[0]])
        df_s_list.append(df_s)
    df_s_all = pd.concat(df_s_list)
    df_s_all.to_excel(writer, sheet_name='Sheet2', startcol=1, index=False)
    workbook  = writer.book
    worksheet1 = writer.sheets['Sheet1']
    worksheet1.set_column('A:A', 30)
    worksheet1.set_column('B:AD', 15)
    worksheet2 = writer.sheets['Sheet2']
    worksheet2.set_column('A:A', 50)
    worksheet2.set_column('B:AD', 15)

    cell_format_bank = workbook.add_format({'bold': False, 'align': 'left', 'num_format': '#,##'})
    cell_format_name = workbook.add_format({'bold': True, 'align': 'left', 'num_format': '#,##'})
    cell_format_name_border = workbook.add_format({'bold': True, 'align': 'left', 'right': 1, 'num_format': '#,##'})
    cell_format_left_bold = workbook.add_format({'bold': False, 'align': 'right', 'border': 1, 'num_format': '#,##'})
    cell_format_border = workbook.add_format({'border': 1, 'num_format': '#,##', 'align': 'right'})
    cell_format_border_green_bold = workbook.add_format({'border': 1, 'num_format': '#,##', 'align': 'right','bg_color': '#ccffcc', 'bold': True})
    cell_format_border_green = workbook.add_format({'border': 1, 'num_format': '#,##', 'align': 'right','bg_color': '#ccffcc'})
    cell_format_border_only = workbook.add_format({'border': 7, 'num_format': '#,##'})
    for i, name in enumerate(banks_data_df.keys()):
        for j, setting in enumerate(settings + growth_settings, 2):
            if setting[0] != highlight_stroke:
                worksheet1.write(i*(len(settings) + len(growth_settings) + 3) + j, 0, f'{setting[0]}', cell_format_left_bold)
                worksheet1.set_row(i*(len(settings) + len(growth_settings) + 3) + j, None, cell_format_border)
                continue
            worksheet1.write(i*(len(settings) + len(growth_settings) + 3) + j, 0, f'{setting[0]}')
            worksheet1.set_row(i*(len(settings) + len(growth_settings) + 3) + j, None, cell_format_border_green_bold)
        worksheet1.write(i*(len(settings) + len(growth_settings) + 3), 0, 'Название банка:', cell_format_bank)
        worksheet1.write(i*(len(settings) + len(growth_settings) + 3), 1, f'{name}', cell_format_name)

    for i, name in enumerate(df_s_all.index.to_list(), 1):
        worksheet2.write(i, 0, f'{name}', cell_format_name_border)
        worksheet2.set_row(i, None, cell_format_border_only)
    workbook.close()

def to_num(settings):
    setting_num = []
    for setting in settings:
        size = len(setting)
        num = []
        number_cells = [2] if size == 3 else [2, 4]
        for cell in number_cells:
            if setting[cell] == 'numsc':
                num.append(1)
            elif setting[cell] == 'vr':
                num.append(2)
            elif setting[cell] == 'vv':
                num.append(3)
            elif setting[cell] == 'vitg':
                num.append(4)
            elif setting[cell] == 'ora':
                num.append(5)
            elif setting[cell] == 'ova':
                num.append(6)
            elif setting[cell] == 'oitga':
                num.append(7)
            elif setting[cell] == 'orp':
                num.append(8)
            elif setting[cell] == 'ovp':
                num.append(9)
            elif setting[cell] == 'oitgp':
                num.append(10)
            elif setting[cell] == 'ir':
                num.append(11)
            elif setting[cell] == 'iv':
                num.append(12)
            elif setting[cell] == 'iitg':
                num.append(13)
            else:
                continue
        if size == 3:
            setting_num.append([setting[0], setting[1], num[0]])
        if size == 5:
            setting_num.append([setting[0], setting[1], num[0], setting[3], num[1]])
    return setting_num

def to_str(settings):
    setting_num = []
    for setting in settings:
        size = len(setting)
        st = []
        number_cells = [2] if size == 3 else [2, 4]
        for cell in number_cells:
            if setting[cell] == 1:
                st.append('numsc')
            elif setting[cell] == 2:
                st.append('vr')
            elif setting[cell] == 3:
                st.append('vv')
            elif setting[cell] == 4:
                st.append('vitg')
            elif setting[cell] == 5:
                st.append('ora')
            elif setting[cell] == 6:
                st.append('ova')
            elif setting[cell] == 7:
                st.append('oitga')
            elif setting[cell] == 8:
                st.append('orp')
            elif setting[cell] == 9:
                st.append('ovp')
            elif setting[cell] == 10:
                st.append('oitgp')
            elif setting[cell] == 11:
                st.append('ir')
            elif setting[cell] == 12:
                st.append('iv')
            elif setting[cell] == 13:
                st.append('iitg')
            else:
                continue
        if size == 3:
            setting_num.append([setting[0], setting[1], st[0]])
        if size == 5:
            setting_num.append([setting[0], setting[1], st[0], setting[3], st[1]])
    return setting_num


settings = [
    ['Выдано БГ', '91315', 'oitgp'],
    ['Погашено БГ', '91315', 'oitga'],
    ['Портфель БГ', '91315', 'iitg'],
    ['Комиссия БГ', '47502', 'oitga'],
    ['Выдачи в рамках лимита', '91319', 'oitga'],
    ['Неис лимиты БГ на дату', '91319', 'iitg'],
    ['Выплачено по требованию БГ', '60315', 'oitga'],
    ['Дефолтность портфеля', '60315', 'oitga', '91315', 'iitg'],
    ['Доходность выдач', '47502', 'oitga', '91315', 'oitgp']
]

growth_settings = [
    ['Изменение портфеля', 3]
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
    ]
]

win1 = sg.Window('101 форма', layout1)
win2_active = False
banks_data, banks_name = {}, {}
dates = []
while True:
    try:
        ev1, val1 = win1.read(timeout=100)
        if ev1 in (None, 'Exit'):
            break
        elif ev1 == 'Взять данные':
            sg.popup_quick_message('Начинаю брать данные', background_color=color_popup, auto_close_duration=2, no_titlebar=True)
            dates = []
            banks_data, banks_name = {}, {}
            t_smonth = temp_smonth
            t_syear = temp_syear
            while True:
                if (t_syear < temp_fyear) or (temp_fyear == t_syear and t_smonth <= temp_fmonth):
                    month = f'{t_smonth:2}'.replace(' ', '0')
                    dates.append(f'{t_syear}-{month}-01')
                    if t_smonth == 12:
                        t_smonth = 1
                        t_syear += 1
                    else:
                        t_smonth += 1
                else:
                    break
            regnums = val1['-REGNUMS-'].replace(', ', ',').replace(' ', ',').split(sep=',')
            if '' in regnums:
                regnums.remove('')
            # проверить на int str
            banks_data, banks_name = get_banks_data_and_name(regnums, settings, dates)
            regnums_for_listbox = [
                f'{code:>8} | {"нет данных за этот период" if isinstance(name, bool) else name:<20}'.replace(' ', ' ')
                for code, name in banks_name.items()
            ]
            win1['-LISTBOX-'].update(regnums_for_listbox)
        elif ev1 == 'Сохранить':
            if banks_data and banks_name:
                save_excel_data(banks_data, banks_name, Path(val1['-PATH-']))
            banks_data, banks_name = {}, {}
            win1['-LISTBOX-'].update([''])
            sg.popup_auto_close('Сохранил файл', background_color=color_popup_ok, auto_close_duration=2, no_titlebar=True)
        elif ev1 == 'Удалить':
            if val1['-LISTBOX-']:
                for key in [value.split()[0] for value in val1['-LISTBOX-']]:
                    try:
                        del banks_data[key]
                        del banks_name[key]
                    except KeyError:
                        continue
            banks_name = {key: banks_name[key] for key in banks_data.keys()}
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
                                        f'{setting[0]:<20} | {setting[1]:>6} | {setting[2]:<3}' if len(setting) == 3
                                        else f'{setting[0]:<20} || {setting[1]:>6} | {setting[2]:<3} / {setting[3]:>6} | {setting[4]:<3}'
                                        for setting in to_num(settings)
                                    ], select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, key='-SLISTBOX-')
                            ],
                            [sg.Button('Вверх'), sg.Button('Вниз'), sg.Button('Удалить')]
                        ]
                    )
                ],
                [
                    sg.Frame('(Имя столбца, номер строки, номер столбца)', [
                            [sg.InputText('', key='-SETTINGS-')],
                            [sg.Button('Добавить')]
                        ]
                    )
                ],
                [
                    sg.Frame('Отслеживание роста', [
                            [
                                sg.Listbox(values=[
                                        f'{sett[0]:<20} | {sett[1]:>3}' for sett in growth_settings
                                    ], select_mode=sg.LISTBOX_SELECT_MODE_MULTIPLE, key='-SLISTBOX_GROWTH-')
                            ],
                            [
                                sg.Button('Удалить отслеживание')
                            ]
                        ]
                    )
                ],
                [
                    sg.Frame('(Имя столбца, номер строки для отслеживания роста)', [
                            [sg.InputText('', key='-SETTINGS2-')],
                            [sg.Button('Добавить отслеживание')]
                        ]
                    )
                ],
                [
                    sg.Frame('Начало', [[
                            sg.Text('Год'), sg.InputText(f'{syear}', size=(45, 20), key='-SYEAR-'),
                            sg.Text('Месяц'), sg.InputText(f'{smonth}', size=(30, 20), key='-SMONTH-'),
                        ]],
                    ),
                    sg.Frame('Конец', [[
                            sg.Text('Год'), sg.InputText(f'{fyear}', size=(45, 20), key='-FYEAR-'),
                            sg.Text('Месяц'), sg.InputText(f'{fmonth}', size=(30, 20), key='-FMONTH-')
                        ]]
                    )
                ],
                [sg.Button('Сохранить даты')],
                [sg.Button('Назад')]
            ]
            win2 = sg.Window('Настройки', layout2)
            def update_setting_listbox():
                lis = [
                    f'{sett[0]:<20} | {sett[1]:>6} | {sett[2]:<3}' if len(sett) == 3
                    else f'{sett[0]:<20} || {sett[1]:>6} | {sett[2]:<3} / {sett[3]:>6} | {sett[4]:<3}'
                    for sett in to_num(settings)
                ]
                win2['-SLISTBOX-'].update(lis)

                lis2 = [
                        f'{sett[0]:<20} | {sett[1]:>3}' for sett in growth_settings
                    ]
                win2['-SLISTBOX_GROWTH-'].update(lis2)


            while True:
                ev2, val2 = win2.Read()
                if ev2 in (None, 'Exit', 'Назад'):
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
                                int(element[1])
                                f_s.append([element[0], element[1], int(element[2])])
                            except (ValueError, IndexError):
                                continue
                        if len(element) == 5:
                            try:
                                int(element[1])
                                int(element[3])
                                f_s.append([element[0], element[1], int(element[2]), element[3], int(element[4])])
                            except (ValueError, IndexError):
                                continue
                    settings += to_str(f_s)
                    update_setting_listbox()
                if ev2 == 'Добавить отслеживание':
                    # filtered settings
                    f_s = []
                    # readed settings
                    r_s = re.split(r'[()]', val2['-SETTINGS2-'])
                    for elem in r_s:
                        element = elem.split(', ')
                        if len(element) == 2:
                            try:
                                int(element[1])
                                f_s.append([element[0], int(element[1])])
                            except (ValueError, IndexError):
                                continue
                    growth_settings += f_s
                    update_setting_listbox()
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
                    update_setting_listbox()
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
                    update_setting_listbox()
                if ev2 == 'Удалить':
                    if val2['-SLISTBOX-']:
                        for key in [value.split('|')[0] for value in val2['-SLISTBOX-']]:
                            zero_set = " ".join([el for el in key.strip().split(' ') if el.strip()])
                            for setting in settings:
                                if setting[0] == zero_set:
                                    settings.remove(setting)
                                    break
                    update_setting_listbox()
                if ev2 == 'Удалить отслеживание':
                    if val2['-SLISTBOX_GROWTH-']:
                        for key in [value.split('|')[0] for value in val2['-SLISTBOX_GROWTH-']]:
                            zero_set = " ".join([el for el in key.strip().split(' ') if el.strip()])
                            for setting in growth_settings:
                                if setting[0] == zero_set:
                                    growth_settings.remove(setting)
                                    break
                    update_setting_listbox()
                if ev2 == 'Сохранить даты':
                    dates = []
                    syear, fyear = int(val2['-SYEAR-']), int(val2['-FYEAR-'])
                    smonth, fmonth = int(val2['-SMONTH-']), int(val2['-FMONTH-'])
                    if smonth == 12:
                        temp_smonth = 1
                        temp_syear = syear + 1
                    else:
                        temp_smonth = smonth + 1
                        temp_syear = syear
                    if fmonth == 12:
                        temp_fmonth = 1
                        temp_fyear = fyear + 1
                    else:
                        temp_fyear = fyear
                        temp_fmonth = fmonth + 1
                    sg.popup_quick_message('Сохранил даты', background_color=color_popup, auto_close_duration=1, no_titlebar=True)
    except Exception as e:
        sg.PopupNonBlocking(e)
        print(e)
win1.close()