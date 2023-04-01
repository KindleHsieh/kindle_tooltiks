# -*- coding: utf-8 -*-
# author: Kindle Hsieh time: 2021/3/16
# import openpyxl as xl
# from openpyxl import Workbook
# from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
# from openpyxl.chart import Reference, LineChart, ScatterChart, Series
# from openpyxl.utils.dataframe import dataframe_to_rows

from xlsxwriter import Workbook

from datetime import datetime

import os
import pandas as pd
import numpy as np
from glob import glob
import logging

#####################################################
"""
4/28
1. 修正了Scales data 只能畫53比的限制，放寬到500比。
2. 修正了Series 資料最後一筆不會畫圖的問題。
3. 將圖片向前移動到最前面。
"""

#####################################################

def set_ordinal_index(df, cols):
    columns, df.columns = df.columns, np.arange(len(df.columns))
    mask = df.columns.isin(cols)
    df = df.set_index(cols)
    df.columns = columns[~mask]
    df.index.names = columns[mask]
    return df

class save_error_log():
    def __init__(self):
        FORMAT = '%(asctime)s %(levelname)s: %(message)s'
        DATE_FORMAT = '%Y%m%d %H:%M:%S'
        logging.basicConfig(level=logging.DEBUG, filename='myLog.log', filemode='w', format=FORMAT, datefmt=DATE_FORMAT)
    def output_log(self):
        logging.debug('debug message', exc_info=True)

def xlsx_column_num2name(num):
    # 26: Z, 27: AA
    # column_name = ''
    if num > 0:
        reminder = (num-1) % 26
        num = (num - reminder) // 26
        return xlsx_column_num2name(num) + chr(65+reminder)
    else:
        return ""



#####################################################
#####################################################

def txt_data(file_name):
    # 使用定義的target list.
    global target, data_sheet, machine_counter, scale_tar_limit_df
    # 獲取現在是哪台機台。
    machine_name = file_name.split(' ')[-1][:-4]
    try:
        machine_counter[machine_name] += 1
    except KeyError:
        machine_counter[machine_name] = 1

    # machine_paras data.
    df = pd.read_csv(file_name)
    # 將第一個的machine_name改成Frequency, 避免後面再更改「第一個參數」欄位名稱的時候出現兩個Frequency。
    df = df.rename(columns={machine_name: 'Frequency'})

    # 參數紀錄。 (init)
    paras_columns = []
    for i, v in enumerate(df.columns):
        if 'Serial' in v:
            paras_columns.append(i)
    paras = {}
    for col in paras_columns:
        # paras = [col, x_axis_type, x_tick, y_tick]
        paras[df.iloc[0, col]] = [col, df.iloc[6, col], df.iloc[10, col], df.iloc[11, col]]
    paras_sorted_list = sorted(list(paras.keys()))

    for tar in target:
        # 第一次讀到這個para的檔案。
        if data_sheet.get(tar) is None:
            # scale Data.
            if df.iloc[:, paras[tar][0] + 2].dropna().shape[0] == 1:
                # Limit Info.
                target_paras_limit = []
                df_ll = pd.DataFrame()
                for par in paras_sorted_list:
                    if (par == tar + ' Upper Limit') or (par == tar + ' Lower Limit'):
                        target_paras_limit.append(par[par.find(tar) + len(tar) + 1:])
                        if 'Phase' in par:
                            if df_ll.empty:
                                df_ll = df.iloc[:, paras[par][0] + 3].dropna()
                            else:
                                df_ll = df_ll.append(df.iloc[:, paras[par][0] + 3].dropna())
                        else:
                            if df_ll.empty:
                                df_ll = df.iloc[:, paras[par][0] + 2].dropna()
                            else:
                                df_ll = df_ll.append(df.iloc[:, paras[par][0] + 2].dropna())
                    elif 'Count' in tar:
                        if (par == tar[:-6] + ' Upper Limit') or (par == tar[:-6] + ' Lower Limit'):
                            target_paras_limit.append(par[par.find(tar[:-6]) + len(tar[:-6]) + 1:])
                            if 'Phase' in par:
                                if df_ll.empty:
                                    df_ll = df.iloc[:, paras[par][0] + 3].dropna()
                                else:
                                    df_ll = df_ll.append(df.iloc[:, paras[par][0] + 3].dropna())
                            else:
                                if df_ll.empty:
                                    df_ll = df.iloc[:, paras[par][0] + 2].dropna()
                                else:
                                    df_ll = df_ll.append(df.iloc[:, paras[par][0] + 2].dropna())
                # 此參數有上下界。
                if len(target_paras_limit) != 0:
                    df_ll = pd.Series(df_ll.values, index=range(1, len(df_ll) + 1))
                    scale_tar_limit_df[tar] = df_ll.rename(machine_name)
                    # 加入data。
                    if 'Phase' in tar:
                        data_sheet[tar] = df.iloc[:, paras[tar][0] + 3].dropna().append(df_ll).rename(machine_name)
                    else:
                        data_sheet[tar] = df.iloc[:, paras[tar][0] + 2].dropna().append(df_ll).rename(machine_name)
                    limit_temp = [paras[tar][3]]
                    limit_temp.extend(target_paras_limit)
                    data_sheet[tar] = pd.concat([pd.Series(limit_temp, name=""), data_sheet[tar]], axis=1)
                # 此參數沒有上下界。
                else:
                    if 'Phase' in tar:
                        data_sheet[tar] = pd.concat([pd.Series([paras[tar][3]], name=""), df.iloc[:, paras[tar][0] + 3].dropna().rename(machine_name)], axis=1)
                    else:
                        data_sheet[tar] = pd.concat([pd.Series([paras[tar][3]], name=""), df.iloc[:, paras[tar][0] + 2].dropna().rename(machine_name)], axis=1)
                # data_sheet[tar].rename(machine_name)
            # Series Data.
            else:
                # Limit Info.
                target_paras_limit = []
                df_ll = pd.DataFrame()
                for par in paras_sorted_list:
                    if (par == tar + ' Upper Limit') or (par == tar + ' Lower Limit'):
                        limit_type = par[par.find(tar) + len(tar) + 1:]
                        target_paras_limit.append(limit_type)
                        if 'Phase' in tar:
                            if df_ll.empty:
                                df_ll = pd.concat([df.iloc[:, paras[par][0] + 3].dropna(),
                                                   df.iloc[:, paras[par][0] + 1].dropna()], axis=1)
                                df_ll = df_ll.rename(
                                    columns=dict(zip(df_ll.columns, [limit_type, 'Frequency'])))
                            else:
                                df_ll = pd.concat([df.iloc[:, paras[par][0] + 3].dropna(), df_ll], axis=1)
                                df_ll = df_ll.rename(columns={df_ll.columns[0]: limit_type})
                        else:
                            if df_ll.empty:
                                df_ll = pd.concat([df.iloc[:, paras[par][0] + 2].dropna(),
                                                   df.iloc[:, paras[par][0] + 1].dropna()], axis=1)
                                df_ll = df_ll.rename(
                                    columns=dict(zip(df_ll.columns, [limit_type, 'Frequency'])))
                            else:
                                df_ll = pd.concat([df.iloc[:, paras[par][0] + 2].dropna(), df_ll], axis=1)
                                df_ll = df_ll.rename(columns={df_ll.columns[0]: limit_type})

                # 此參數有上下界。
                if len(target_paras_limit) != 0:
                    # 加入data.
                    if 'Phase' in tar:
                        df_data_temp = pd.concat([df.iloc[:, paras[tar][0] + 1].dropna(), df.iloc[:, paras[tar][0] + 3].dropna().rename(machine_name)], axis=1)
                    else:
                        df_data_temp = pd.concat([df.iloc[:, paras[tar][0] + 1].dropna(), df.iloc[:, paras[tar][0] + 2].dropna().rename(machine_name)], axis=1)
                    # df_data_temp = df.iloc[:, paras[tar][0] + 1:paras[tar][0] + 3]
                    df_temp = df_ll.merge(df_data_temp, left_on='Frequency', right_on=df_data_temp.columns[0], how='outer', sort=True)
                    # df_temp = pd.concat([df_ll, df.iloc[:, paras[tar][0] + 1].dropna(), df.iloc[:, paras[tar][0] + 2].dropna()], axis=1)
                    if len(target_paras_limit) == 1:
                        df_temp = df_temp.rename(columns={df_temp.columns[2]: 'Frequency'})
                    elif len(target_paras_limit) == 2:
                        df_temp = df_temp.rename(columns={df_temp.columns[3]: 'Frequency'})
                    data_sheet[tar] = df_temp
                # 此參數沒有上下界。
                else:
                    if 'Phase' in tar:
                        df_temp = pd.concat([df.iloc[:, paras[tar][0] + 1].dropna(), df.iloc[:, paras[tar][0] + 3].dropna().rename(machine_name)], axis=1)
                    else:
                        df_temp = pd.concat([df.iloc[:, paras[tar][0] + 1].dropna(), df.iloc[:, paras[tar][0] + 2].dropna().rename(machine_name)], axis=1)
                    df_temp = df_temp.rename(columns={df_temp.columns[0]: 'Frequency'})
                    data_sheet[tar] = df_temp
            # Counter.
            data_y_count[tar] = 1
        # 以後讀到這個para的檔案。
        else:
            if 'Phase' in tar:
                data_column = 3
            else:
                data_column = 2
            # scale Data.
            if df.iloc[:, paras[tar][0] + data_column].dropna().shape[0] == 1:
                # 加入data。
                if scale_tar_limit_df.get(tar) is None:
                    df_ = df.iloc[:, paras[tar][0] + data_column].dropna().rename(machine_name)
                else:
                    df_ = df.iloc[:, paras[tar][0] + data_column].dropna().append(scale_tar_limit_df[tar]).rename(machine_name)
                # df_.rename(machine_name)
                data_sheet[tar] = pd.concat([data_sheet[tar], df_], axis=1)
            # Series Data.
            else:
                # 加入data.
                freq_pos = [pos for pos, col in enumerate(data_sheet[tar].columns) if col == 'Frequency']
                # data_sheet[tar] = pd.concat([data_sheet[tar], df.iloc[:, paras[tar][0] + data_column].dropna().rename(machine_name)], axis=1)
                df_ = df.iloc[:, [paras[tar][0] + 1, paras[tar][0] + data_column]].dropna()
                df__ = data_sheet[tar].iloc[:, [freq_pos[-1]]]
                df__ = df__.join(df_.set_index(df_.columns[0]), on='Frequency')[df_.columns[-1]]
                data_sheet[tar] = pd.concat([data_sheet[tar], df__.rename(machine_name)], axis=1)
            # Counter.
            data_y_count[tar] += 1
    return paras


def create_final_report(data_sheet, scale_tar_limit_df, paras_info, out_put_file_name):
    # 為了能照著輸入進來的順序做標籤。
    global target
    # 開始創造excel.
    wb_chart = Workbook(out_put_file_name)
    limit_format = wb_chart.add_format({'bold': True, 'font_color': 'red'})
    fail_format = wb_chart.add_format({'bg_color': '#F4B9C1'})

    for sheet_par in target:
        # 創建sheet.
        ws = wb_chart.add_worksheet(sheet_par)
        # 分為 Series data and Scale data.
        freq_pos = []
        if (type(data_sheet[sheet_par]) is pd.DataFrame) or (isinstance(data_sheet[sheet_par], pd.DataFrame)):
            freq_pos = [pos for pos, col in enumerate(data_sheet[sheet_par].columns) if col == 'Frequency']
            column_names = [col[:col.find('.')] if col.find('.') > -1 else col for col in data_sheet[sheet_par].columns]

        # Scale data.
        if len(freq_pos) == 0:
            # -Machine Name.
            if (type(data_sheet[sheet_par]) is pd.DataFrame) or (isinstance(data_sheet[sheet_par], pd.DataFrame)):
                ws.write_row(f'A3', column_names)
            # 當只有一筆資料的時候。
            else:
                ws.write(f'B3', data_sheet[sheet_par].name)
                # Data.
                np_data = data_sheet[sheet_par].to_numpy().T
                # -Pass or FAIL and Series Num.
                measure_counter = 1
                ws.write(f'{chr(66)}1', measure_counter)
                # 有Limit.
                if not scale_tar_limit_df.get(sheet_par) is None:
                    if len(np_data) == 3:
                        if (np_data[0] < np_data[1]) or (np_data[0] > np_data[2]):
                            ws.write(f'{chr(66)}2', "FAIL", limit_format)
                        else:
                            ws.write(f'{chr(66)}2', "PASS")
                    else:
                        if ws[f'{chr(65)}5'] == 'Upper Limit':
                            if np_data[0] > np_data[1]:
                                ws.write(f'{chr(66)}2', "FAIL", limit_format)
                            else:
                                ws.write(f'{chr(66)}2', "PASS")

                        elif ws[f'{chr(65)}5'] == 'Lower Limit':
                            if np_data[0] < np_data[1]:
                                ws.write(f'{chr(66)}2', "PASS")
                            else:
                                ws.write(f'{chr(66)}2', "FAIL", limit_format)
                # -data.
                ws.write_column(f'{chr(65)}4', np_data)
                continue
            # Data.
            np_data = data_sheet[sheet_par].to_numpy().T
            measure_counter = 0
            if np_data.ndim == 2:
                for col, data in enumerate(np_data, 1):
                    # 放Scale的極限值資訊。
                    if col == 1:
                        ws.write_column(f'{xlsx_column_num2name(col)}4', data)
                        continue
                    # -PASS or FAIL and Series Num.
                    # 第幾筆資料。
                    # Add Counter.
                    measure_counter += 1
                    ws.write(f'{xlsx_column_num2name(col)}1', measure_counter)
                    # 有Limit.
                    if not scale_tar_limit_df.get(sheet_par) is None:
                        if len(data) == 3:
                            if (data[0] < data[1]) or (data[0] > data[2]):
                                ws.write(f'{xlsx_column_num2name(col)}2', "FAIL", limit_format)
                            else:
                                ws.write(f'{xlsx_column_num2name(col)}2', "PASS")
                        else:
                            if np_data[0][1] == 'Upper Limit':
                                if data[0] > data[1]:
                                    ws.write(f'{xlsx_column_num2name(col)}2', "FAIL", limit_format)
                                else:
                                    ws.write(f'{xlsx_column_num2name(col)}2', "PASS")

                            elif np_data[0][1] == 'Lower Limit':
                                if data[0] < data[1]:
                                    ws.write(f'{xlsx_column_num2name(col)}2', "PASS")
                                else:
                                    ws.write(f'{xlsx_column_num2name(col)}2', "FAIL", limit_format)

                    # -data.
                    # if not type(data) is np.float64:
                    #     if sum(np.isnan(data)) > 0:
                    #         data = map(str, data)
                    #     ws.write_column(f'{chr(65+col)}4', data)
                    # else:
                    #     ws.write(f'{chr(65+col)}4', data)
                    if len(data) == 1:
                        ws.write_column(f'{xlsx_column_num2name(col)}4', data)
                    else:
                        ws.write(f'{xlsx_column_num2name(col)}4', data[0])
                        ws.write_column(f'{xlsx_column_num2name(col)}5', data[1:], limit_format)

                # Plotting.
                chart_ = wb_chart.add_chart({'type': 'scatter'})
                # 有Limit.
                if not scale_tar_limit_df.get(sheet_par) is None:
                    if np_data.shape[1] == 3:
                        # Plotting.
                        chart_ = wb_chart.add_chart({'type': 'scatter'})
                        # 前兩條是 Upper and Lower.
                        chart_.add_series({
                            'name': f"='{sheet_par}'!$A$5",
                            'categories': f"='{sheet_par}'!$B$1:${xlsx_column_num2name(500)}$1",
                            'values': f"='{sheet_par}'!$B$5:${xlsx_column_num2name(500)}$5",
                            'line': {'color': 'red'},
                            'marker': {'type': 'none'}
                        })
                        chart_.add_series({
                            'name': f"='{sheet_par}'!$A$6",
                            'categories': f"='{sheet_par}'!$B$1:${xlsx_column_num2name(500)}$1",
                            'values': f"='{sheet_par}'!$B$6:${xlsx_column_num2name(500)}$6",
                            'line': {'color': 'red'},
                            'marker': {'type': 'none'}
                        })
                        # 後面是畫資料.
                        chart_.add_series({
                            'name': f"='{sheet_par}'!$A$4",
                            'categories': f"='{sheet_par}'!$B$1:${xlsx_column_num2name(500)}$1",
                            'values': f"='{sheet_par}'!$B$4:${xlsx_column_num2name(500)}$4",
                            # 'line': {'width': 1}
                        })
                    # 只有一個limit.
                    elif np_data.shape[1] == 2:
                        # Upper or Lower.
                        chart_.add_series({
                            'name': f"='{sheet_par}'!$A$5",
                            'categories': f"='{sheet_par}'!$B$1:${xlsx_column_num2name(500)}$1",
                            'values': f"='{sheet_par}'!$B$5:${xlsx_column_num2name(500)}$5",
                            'line': {'color': 'red'},
                            'marker': {'type': 'none'}
                        })
                        # 後面是畫資料.
                        chart_.add_series({
                            'name': f"='{sheet_par}'!$A$4",
                            'categories': f"='{sheet_par}'!$B$1:${xlsx_column_num2name(500)}$1",
                            'values': f"='{sheet_par}'!$B$4:${xlsx_column_num2name(500)}$4",
                            # 'line': {'width': 1}
                        })
                # 沒Limit.
                else:
                    # 後面是畫資料.
                    chart_.add_series({
                        'name': f"='{sheet_par}'!$A$4",
                        'categories': f"='{sheet_par}'!$B$1:${xlsx_column_num2name(500)}$1",
                        'values': f"='{sheet_par}'!$B$4:${xlsx_column_num2name(500)}$4"
                    })

                # # Set an Excel chart style.
                chart_.set_style(11)

                # Add a chart title and some axis labels.
                chart_.set_size({'width': 720, 'height': 576})
                chart_.set_title({'name': f'1-{measure_counter} units {sheet_par}'})
                chart_.set_x_axis({'name': 'Samples', 'label_position': 'low', 'major_unit': 1})
                chart_.set_y_axis({'name': paras_info[sheet_par][3]})

                # Insert the chart into the worksheet (with an offset).
                ws.insert_chart(f'B9', chart_)

        # Series data.
        else:
            # 圖片在前面，Table在後面： Tabel從 R(18) 開始。
            shift_columns = 17

            # Data.
            # -Machine Name.
            ws.write_row(f'{xlsx_column_num2name(1+shift_columns)}3', column_names)

            np_data = data_sheet[sheet_par].to_numpy().T
            print(sheet_par)
            measure_counter = 0
            for col, data in enumerate(np_data, 1):
                # -Pass or FAIL and Series Num.
                if len(freq_pos) == 2 and col-1 > freq_pos[-1]:
                    # 第幾筆資料。
                    # Add Counter.
                    measure_counter += 1
                    ws.write(f'{xlsx_column_num2name(col+shift_columns)}1', measure_counter)
                    # PASS or FAIL字。
                    if freq_pos[0] == 2:  # 有 Upper and Lower.
                        # - 找出 limit 關心的數值。 分為 Upper 和 Lower.
                        # - Filter.
                        # Upper 不nan.
                        upper_nonan = np.where(~np.isnan(np_data[1]))
                        # Lower 不nan.
                        lower_nonan = np.where(~np.isnan(np_data[0]))
                        greater_filter = np.greater(data[upper_nonan], np_data[0][upper_nonan])
                        lesser_filter = np.less(data[lower_nonan], np_data[1][lower_nonan])
                        # 找有沒有超出界線的數值。
                        PF_filter = greater_filter | lesser_filter
                        if sum(PF_filter) > 0:
                            ws.write(f'{xlsx_column_num2name(col+shift_columns)}2', "FAIL", limit_format)
                        else:
                            ws.write(f'{xlsx_column_num2name(col+shift_columns)}2', "PASS")
                    elif freq_pos[0] == 1:  # 有 Upper or Lower.
                        _nonan = np.where(~np.isnan(np_data[0]))
                        if column_names[0] == 'Upper Limit':
                            # Upper 不nan.
                            PF_filter = np.greater(data[_nonan], np_data[0][_nonan])
                        elif column_names[0] == 'Lower Limit':
                            # Lower 不nan.
                            PF_filter = np.less(data[_nonan], np_data[0][_nonan])
                        if sum(PF_filter) > 0:
                            ws.write(f'{xlsx_column_num2name(col+shift_columns)}2', "FAIL", limit_format)
                        else:
                            ws.write(f'{xlsx_column_num2name(col+shift_columns)}2', "PASS")

                # No Limit data.
                elif len(freq_pos) == 1 and col-1 > freq_pos[-1]:
                    # 第幾筆資料。
                    # Add Counter.
                    measure_counter += 1
                    ws.write(f'{xlsx_column_num2name(col+shift_columns)}1', measure_counter)


                # -data.
                if not type(data) is np.float64:
                    # 有nan的情形處理方式。
                    if sum(np.isnan(data)) > 0:
                        if col-1 < freq_pos[0]:
                            for i, n in enumerate(zip(data, np.isnan(data))):
                                if not n[1]:
                                    ws.write_number(f'{xlsx_column_num2name(col+shift_columns)}{4+i}', n[0], limit_format)
                        else:
                            for i, n in enumerate(zip(data, np.isnan(data))):
                                if not n[1]:
                                    ws.write_number(f'{xlsx_column_num2name(col+shift_columns)}{4+i}', n[0])
                    else:
                        # 沒有nan的處理。
                        if col-1 < freq_pos[0]:
                            ws.write_column(f'{xlsx_column_num2name(col+shift_columns)}4', data, limit_format)
                        else:
                            ws.write_column(f'{xlsx_column_num2name(col+shift_columns)}4', data)
                else:
                    ws.write(f'{xlsx_column_num2name(col+shift_columns)}4', data)
            # 資料都輸入後，開始畫圖跟標記FAIL的顏色。
            # 有 Limit data.
            if len(freq_pos) == 2:
                if freq_pos[0] == 2:  # 有 Upper and Lower.
                    for i in np.where(~np.isnan(np_data[freq_pos[0]]))[0]:
                        upper_pointer = f"{xlsx_column_num2name(1+shift_columns)}${i+4}"
                        lower_pointer = f"{xlsx_column_num2name(2+shift_columns)}${i+4}"
                        for j in range(freq_pos[-1]+1, len(np_data)):
                            data_pointer = f"{xlsx_column_num2name(j+1+shift_columns)}${i+4}"
                            ws.conditional_format(data_pointer,
                                                  {
                                                      'type': 'cell',
                                                      'criteria': ">",
                                                      'value': upper_pointer,
                                                      'format': fail_format
                                                  })
                            ws.conditional_format(data_pointer,
                                                  {
                                                      'type': 'cell',
                                                      'criteria': "<",
                                                      'value': lower_pointer,
                                                      'format': fail_format
                                                  })
                    # Plotting.
                    chart_ = wb_chart.add_chart({'type': 'scatter', 'subtype': 'straight'})
                    # 前兩條是 Upper and Lower.
                    chart_.add_series({
                        'name': f"='{sheet_par}'!${xlsx_column_num2name(1+shift_columns)}$3",
                        'categories': f"='{sheet_par}'!${xlsx_column_num2name(3+shift_columns)}$4:${xlsx_column_num2name(3+shift_columns)}$500",
                        'values': f"='{sheet_par}'!${xlsx_column_num2name(1+shift_columns)}$4:${xlsx_column_num2name(1+shift_columns)}$500",
                        'line': {'color': 'red'}
                    })
                    chart_.add_series({
                        'name': f"='{sheet_par}'!${xlsx_column_num2name(2+shift_columns)}$3",
                        'categories': f"='{sheet_par}'!${xlsx_column_num2name(3+shift_columns)}$4:${xlsx_column_num2name(3+shift_columns)}$500",
                        'values': f"='{sheet_par}'!${xlsx_column_num2name(2+shift_columns)}$4:${xlsx_column_num2name(2+shift_columns)}$500",
                        'line': {'color': 'red'}
                    })
                    # 後面是畫資料.
                    for j in range(freq_pos[-1] + 1, len(np_data)):
                        # print(f"='{sheet_par}'!${chr(65+j)}$4:${chr(65+j)}$500")
                        chart_.add_series({
                            'name': f"='{sheet_par}'!${xlsx_column_num2name(j+1+shift_columns)}$3",
                            'categories': f"='{sheet_par}'!${xlsx_column_num2name(4+shift_columns)}$4:${xlsx_column_num2name(4+shift_columns)}$500",
                            'values': f"='{sheet_par}'!${xlsx_column_num2name(j+1+shift_columns)}$4:${xlsx_column_num2name(j+1+shift_columns)}$500",
                            'line': {'width': 1}
                        })

                elif freq_pos[0] == 1:  # 有 Upper or Lower.
                    if column_names[0] == 'Upper Limit':
                        for i in np.where(~np.isnan(np_data[freq_pos[0]]))[0]:
                            upper_pointer = f"{xlsx_column_num2name(1+shift_columns)}${i+4}"
                            for j in range(freq_pos[-1] + 1, len(np_data)):
                                data_pointer = f"{xlsx_column_num2name(j+1+shift_columns)}${i+4}"
                                ws.conditional_format(data_pointer,
                                                      {
                                                          'type': 'cell',
                                                          'criteria': ">",
                                                          'value': upper_pointer,
                                                          'format': fail_format
                                                      })
                    elif column_names[0] == 'Lower Limit':
                        for i in np.where(~np.isnan(np_data[freq_pos[0]]))[0]:
                            lower_pointer = f"{xlsx_column_num2name(1+shift_columns)}${i+4}"
                            for j in range(freq_pos[-1] + 1, len(np_data)):
                                data_pointer = f"{xlsx_column_num2name(j+1+shift_columns)}${i+4}"
                                ws.conditional_format(data_pointer,
                                                      {
                                                          'type': 'cell',
                                                          'criteria': "<",
                                                          'value': lower_pointer,
                                                          'format': fail_format
                                                      })
                    # Plotting.
                    chart_ = wb_chart.add_chart({'type': 'scatter', 'subtype': 'straight'})
                    chart_.add_series({
                        'name': f"='{sheet_par}'!${xlsx_column_num2name(1+shift_columns)}$3",
                        'categories': f"='{sheet_par}'!${xlsx_column_num2name(2+shift_columns)}$4:${xlsx_column_num2name(2+shift_columns)}$500",
                        'values': f"='{sheet_par}'!${xlsx_column_num2name(1+shift_columns)}$4:${xlsx_column_num2name(1+shift_columns)}$500",
                        'line': {'color': 'red'}
                    })
                    # 後面是畫資料.
                    for j in range(freq_pos[-1] + 1, len(np_data)):
                        # print(f"='{sheet_par}'!${chr(65+j)}$4:${chr(65+j)}$500")
                        chart_.add_series({
                            'name': f"='{sheet_par}'!${xlsx_column_num2name(j+1+shift_columns)}$3",
                            'categories': f"='{sheet_par}'!${xlsx_column_num2name(3+shift_columns)}$4:${xlsx_column_num2name(3+shift_columns)}$500",
                            'values': f"='{sheet_par}'!${xlsx_column_num2name(j+1+shift_columns)}$4:${xlsx_column_num2name(j+1+shift_columns)}$500",
                            'line': {'width': 1}
                        })

                # # Add a chart title and some axis labels.
                # chart_.set_size({'width': 720, 'height': 576})
                # chart_.set_title({'name': f'1-{measure_counter} units {sheet_par}'})
                # chart_.set_x_axis({'name': 'Frequency(Hz)', 'log_base': 10, 'label_position': 'low',
                #                    'major_gridlines': {'visible': True, 'line': {'width': 1.25}},
                #                    'minor_gridlines': {'visible': True, 'line': {'width': 1.25}},
                #                    'min': 10, 'max': 20000
                #                    })
                # chart_.set_y_axis({'name': paras_info[sheet_par][3]})
                #
                # # Set an Excel chart style.
                # chart_.set_style(10)
                #
                # # Insert the chart into the worksheet (with an offset).
                # ws.insert_chart(f'{chr(65+len(np_data))}3', chart_)

            # 沒有 Limit data.
            elif len(freq_pos) == 1:
                # Plotting.
                chart_ = wb_chart.add_chart({'type': 'scatter', 'subtype': 'straight'})

                # 後面是畫資料.
                for j in range(freq_pos[-1] + 1, len(np_data)):
                    chart_.add_series({
                        'name': f"='{sheet_par}'!${xlsx_column_num2name(j+1+shift_columns)}$3",
                        'categories': f"='{sheet_par}'!${xlsx_column_num2name(1+shift_columns)}$4:${xlsx_column_num2name(1+shift_columns)}$500",
                        'values': f"='{sheet_par}'!${xlsx_column_num2name(j+1+shift_columns)}$4:${xlsx_column_num2name(j+1+shift_columns)}$500",
                        'line': {'width': 1}
                    })
            # Add a chart title and some axis labels.
            chart_.set_size({'width': 1020, 'height': 576})
            chart_.set_title({'name': f'1-{measure_counter} units {sheet_par}'})
            chart_.set_x_axis({'name': 'Frequency(Hz)', 'log_base': 10, 'label_position': 'low',
                               'major_gridlines': {'visible': True, 'line': {'width': 1.25}},
                               'minor_gridlines': {'visible': True, 'line': {'width': 1.25}},
                               'min': np.nanmin(np_data[freq_pos]), 'max': np.nanmax(np_data[freq_pos])
                               })
            chart_.set_y_axis({'name': paras_info[sheet_par][3]})

            # Set an Excel chart style.
            chart_.set_style(10)

            # Insert the chart into the worksheet (with an offset).
            ws.insert_chart(f'{xlsx_column_num2name(1)}3', chart_)
    wb_chart.close()
    # 分為limit－ line 及 dot 的類型。 用column名稱是不是有0 1的名稱。
    # Dot.
    # 加入資料 第幾筆數 及 有沒有通過 ，並將 未通過的數據標示紅底 。
    # Limit值是用紅色字，
    # 根據資料畫圖。
    # line.
    # 尋找Frequency的位置：有兩種情形：1. 只有一個freq，也就只有測量資料. 2. 有兩個freq，有limit存在。
    # 加入資料 第幾筆數 及 有沒有通過 ，並將 未通過的數據標示紅底 。
    # Limit值是用紅色字，
    # 根據資料畫圖。


if __name__ == '__main__':

    LOG = save_error_log()
    try:
        with open('target_name.txt', 'r') as f:
            content = f.read()
        target = content.split('\n')
        # target = ['L_FR', 'L_HOHD', 'L_HOHD-2', 'L_LPA Count', 'L_PRB', 'L_THD', 'Mic Sensitivity Avg', 'Mic1_FR', 'Mic1_FR Aligned', 'Mic1_FR_wo offset', 'Mic1_Loopback_L_FR', 'Mic1_Loopback_L_HOHD', 'Mic1_Loopback_L_PRB', 'Mic1_Loopback_L_THD', 'Mic1_Loopback_R_FR', 'Mic1_Loopback_R_HOHD', 'Mic1_Loopback_R_PRB', 'Mic1_Loopback_R_THD', 'Mic1_Loopback_WF_FR', 'Mic1_Loopback_WF_HOHD', 'Mic1_Loopback_WF_PRB', 'Mic1_Loopback_WF_THD', 'Mic1_Mic2_Duplicate', 'Mic1_Seal_Delta', 'Mic1_Sensitivity', 'Mic1_Sensitivity_Delta', 'Mic1_THD', 'Mic1_blocked_FR', 'Mic2_FR', 'Mic2_FR Aligned', 'Mic2_FR_wo offset', 'Mic2_Loopback_L_FR', 'Mic2_Loopback_L_HOHD', 'Mic2_Loopback_L_PRB', 'Mic2_Loopback_L_THD', 'Mic2_Loopback_R_FR', 'Mic2_Loopback_R_HOHD', 'Mic2_Loopback_R_PRB', 'Mic2_Loopback_R_THD', 'Mic2_Loopback_WF_FR', 'Mic2_Loopback_WF_HOHD', 'Mic2_Loopback_WF_PRB', 'Mic2_Loopback_WF_THD', 'Mic2_Relative Phase', 'Mic2_Seal_Delta', 'Mic2_Sensitivity', 'Mic2_Sensitivity_Delta', 'Mic2_THD', 'Mic2_blocked_FR', 'Mic3_FR', 'Mic3_FR Aligned', 'Mic3_FR_wo offset', 'Mic3_Loopback_L_FR', 'Mic3_Loopback_L_HOHD', 'Mic3_Loopback_L_PRB', 'Mic3_Loopback_L_THD', 'Mic3_Loopback_R_FR', 'Mic3_Loopback_R_HOHD', 'Mic3_Loopback_R_PRB', 'Mic3_Loopback_R_THD', 'Mic3_Loopback_WF_FR', 'Mic3_Loopback_WF_HOHD', 'Mic3_Loopback_WF_PRB', 'Mic3_Loopback_WF_THD', 'Mic3_Mic4_Duplicate', 'Mic3_Relative Phase', 'Mic3_Seal_Delta', 'Mic3_Sensitivity', 'Mic3_Sensitivity_Delta', 'Mic3_THD', 'Mic3_blocked_FR', 'Mic4_FR', 'Mic4_FR Aligned', 'Mic4_FR_wo offset', 'Mic4_Loopback_L_FR', 'Mic4_Loopback_L_HOHD', 'Mic4_Loopback_L_PRB', 'Mic4_Loopback_L_THD', 'Mic4_Loopback_R_FR', 'Mic4_Loopback_R_HOHD', 'Mic4_Loopback_R_PRB', 'Mic4_Loopback_R_THD', 'Mic4_Loopback_WF_FR', 'Mic4_Loopback_WF_HOHD', 'Mic4_Loopback_WF_PRB', 'Mic4_Loopback_WF_THD', 'Mic4_Relative Phase', 'Mic4_Seal_Delta', 'Mic4_Sensitivity', 'Mic4_Sensitivity_Delta', 'Mic4_THD', 'Mic4_blocked_FR', 'Prevent data shift', 'R_FR', 'R_HOHD', 'R_HOHD-2', 'R_LPA Count', 'R_PRB', 'R_THD', 'WF_FR', 'WF_HOHD', 'WF_HOHD-2', 'WF_LPA Count', 'WF_PRB', 'WF_THD']
        # target = ['L_FR', 'L_HOHD', 'L_HOHD-2', 'L_LPA Count', 'L_PRB', 'L_THD', 'Mic Sensitivity Avg', 'Mic1_FR', 'Mic1_FR Aligned', 'Mic1_FR_wo offset', 'Mic1_Loopback_L_FR', 'Mic1_Loopback_L_HOHD', 'Mic1_Loopback_L_PRB', 'Mic1_Loopback_L_THD', 'Mic1_Loopback_R_FR', 'Mic1_Loopback_R_HOHD', 'Mic1_Loopback_R_PRB', 'Mic1_Loopback_R_THD', 'Mic1_Loopback_WF_FR', 'Mic1_Loopback_WF_HOHD', 'Mic1_Loopback_WF_PRB', 'Mic1_Loopback_WF_THD', 'Mic1_Mic2_Duplicate', 'Mic1_Seal_Delta', 'Mic1_Sensitivity', 'Mic1_Sensitivity_Delta', 'Mic1_THD', 'Mic1_blocked_FR', 'Mic2_FR', 'Mic2_FR Aligned', 'Mic2_FR_wo offset', 'Mic2_Loopback_L_FR', 'Mic2_Loopback_L_HOHD', 'Mic2_Loopback_L_PRB', 'Mic2_Loopback_L_THD', 'Mic2_Loopback_R_FR', 'Mic2_Loopback_R_HOHD', 'Mic2_Loopback_R_PRB', 'Mic2_Loopback_R_THD', 'Mic2_Loopback_WF_FR', 'Mic2_Loopback_WF_HOHD', 'Mic2_Loopback_WF_PRB', 'Mic2_Loopback_WF_THD', 'Mic2_Relative Phase', 'Mic2_Seal_Delta', 'Mic2_Sensitivity', 'Mic2_Sensitivity_Delta', 'Mic2_THD', 'Mic2_blocked_FR', 'Mic3_FR', 'Mic3_FR Aligned', 'Mic3_FR_wo offset', 'Mic3_Loopback_L_FR', 'Mic3_Loopback_L_HOHD', 'Mic3_Loopback_L_PRB', 'Mic3_Loopback_L_THD', 'Mic3_Loopback_R_FR', 'Mic3_Loopback_R_HOHD', 'Mic3_Loopback_R_PRB', 'Mic3_Loopback_R_THD', 'Mic3_Loopback_WF_FR', 'Mic3_Loopback_WF_HOHD', 'Mic3_Loopback_WF_PRB', 'Mic3_Loopback_WF_THD', 'Mic3_Mic4_Duplicate', 'Mic3_Relative Phase', 'Mic3_Seal_Delta', 'Mic3_Sensitivity', 'Mic3_Sensitivity_Delta', 'Mic3_THD', 'Mic3_blocked_FR', 'Mic4_FR', 'Mic4_FR_wo offset', 'Mic4_Loopback_L_FR', 'Mic4_Loopback_L_HOHD', 'Mic4_Loopback_L_PRB', 'Mic4_Loopback_L_THD', 'Mic4_Loopback_R_FR', 'Mic4_Loopback_R_HOHD', 'Mic4_Loopback_R_PRB', 'Mic4_Loopback_R_THD', 'Mic4_Loopback_WF_FR', 'Mic4_Loopback_WF_HOHD', 'Mic4_Loopback_WF_PRB', 'Mic4_Loopback_WF_THD', 'Mic4_Relative Phase', 'Mic4_Seal_Delta', 'Mic4_Sensitivity', 'Mic4_Sensitivity_Delta', 'Mic4_THD', 'Mic4_blocked_FR', 'Prevent data shift', 'R_FR', 'R_HOHD', 'R_HOHD-2', 'R_LPA Count', 'R_PRB', 'R_THD', 'WF_FR', 'WF_HOHD', 'WF_HOHD-2', 'WF_LPA Count', 'WF_PRB', 'WF_THD']
        # target = ['Mic2_FR Aligned', 'Mic2_Relative Phase']
        # target = ['Mic2_FR Aligned', 'L_FR']
        # target = ['L_FR']
        # target = ['Mic4_FR Aligned']
        # target = ['Mic2_Seal_Delta', 'Mic3_Seal_Delta', 'Prevent data shift', 'Mic2_FR Aligned', 'Mic2_Relative Phase', 'L_FR']
        # target = ['Prevent data shift']
        # target = ['WF_LPA Count', 'R_LPA Count', 'Mic4_Sensitivity_Delta']
        target = ['LPR_FR', 'RPR_FR']
        data_sheet = {}
        data_y_count = {}
        machine_counter = {}
        scale_tar_limit_df = {}
        # file_name = '2021-01-25 114544 821PGC0160111Y19.txt'
        file_name = []
        for folder_content in os.listdir('measure_files_txt'):
            if os.path.isdir(os.path.join('measure_files_txt', folder_content)):
                txt = sorted(glob(fr'measure_files_txt/{folder_content}/*.txt'), key=len, reverse=True)[0]
                file_name.append(txt)

        for fn in file_name:
            print(fn)
            paras_info = txt_data(fn)
        create_final_report(data_sheet, scale_tar_limit_df, paras_info, f"result_test_{datetime.now().replace(microsecond=0).strftime('%Y%m%d%H%M%S')}.xlsx")
    except:
        LOG.output_log()
