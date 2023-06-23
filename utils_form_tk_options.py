import pandas as pd
import numpy as np
import os, sys, glob
import humanize
import re
import xlrd

import json
import itertools
#from urllib.request import urlopen
#import requests, xmltodict
import time, datetime
import math
from pprint import pprint
import gc
from tqdm import tqdm
tqdm.pandas()
import pickle

import logging
import zipfile
import warnings
import argparse

import warnings
warnings.filterwarnings("ignore")

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.utils import units
from openpyxl.styles import Border, Side, PatternFill, GradientFill, Alignment
from openpyxl import drawing

import matplotlib.pyplot as plt
# import seaborn as sns
# %matplotlib inline
from matplotlib.colors import ListedColormap, BoundaryNorm

from utils_form_tk_enrichment_options import preprocess_tkbd_options
from utils_io import save_df_lst_to_excel
from utils_io import logger

if len(logger.handlers) > 1:
    for handler in logger.handlers:
        logger.removeHandler(handler)
    from utils_io import logger

from  local_dictionaries import sevice_sections, service_chapters, service_types_A, service_types_B, service_classes_A, service_classes_B
from  local_dictionaries import dict_ath_anatomy, dict_ath_therapy, dict_ath_pharm, dict_ath_chemical

def read_enriched_tk_data(path_tkbd_processed, fn_tk_bd):
    df_services = pd.read_excel(os.path.join(path_tkbd_processed, fn_tk_bd), sheet_name = 'Услуги')
    print(df_services.shape)
    display(df_services.head(2))
    df_LP = pd.read_excel(os.path.join(path_tkbd_processed, fn_tk_bd), sheet_name = 'ЛП')
    print(df_desc.shape)
    display(df_LP.head(2))
    df_RM = pd.read_excel(os.path.join(path_tkbd_processed, fn_tk_bd), sheet_name = 'РМ')
    print(df_RM.shape)
    display(df_RM.head(2))
    return df_services, df_LP, df_RM




def create_tk_models_dict_02(model, xls_file, tk_code=7777777, tk_name='tk_test', profile='profile_test', tk_models = {} ):
    # tk_models = {}
    max_len = 40

    if tk_models.get(tk_name) is None:
        tk_models[tk_name] = {}
    tk_models[tk_name]['Код ТК'] = tk_code
    tk_models[tk_name]['Профиль'] = profile
    tk_models[tk_name]['Наименование ТК (короткое)'] = tk_name[:max_len]
    # 'Модели': [{'Модель пациента': 'Техно',
    #             'Файл Excel': 'Аденома_предстательной_железы_Склероз_шейки_мочевого_пузыря_Стриктура '
    #                           'техно.xlsx'},
    #           {'Модель пациента': 'База',
    #             'Файл Excel': 'Аденома_предстательной_железы_Склероз_шейки_мочевого_пузыря_Стриктура.xlsx'}],

    # tk_models[tk_name]['Модели'] = [models]
        # tk_models.setdefault('tk_name', []).append(row['Наименование ТК'])
    #tk_models[tk_name]['Модели'].append (dict(zip(['Модель пациента', 'Файл Excel',
    #  '#Название листа в файле Excel', 'Услуги', 'ЛП', 'РМ'], row.values[4:])))
    # tk_models[tk_name]['Модели'].append (dict(zip(['Модель пациента', 'Файл Excel',], row.values[4:6])))
    tk_models[tk_name]['Модели'] = []
    for i_m, model in enumerate(models):
        tk_models[tk_name]['Модели'].append (dict(zip(['Модель пациента', 'Файл Excel',], [model, xls_files[i_m]])))

    return tk_models

def format_excel_cols_short(ws, format_cols, auto_filter=False):
    l_alignment=Alignment(horizontal='left', vertical= 'top', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
    r_alignment=Alignment(horizontal='right', vertical= 'top', text_rotation=0, wrap_text=True, shrink_to_fit=False, indent=0)
    last_cell = ws.cell(row=1, column=len(format_cols))
    full_range = "A1:" + last_cell.column_letter + str(ws.max_row)
    if auto_filter:
        ws.auto_filter.ref = full_range
    ws.freeze_panes = ws['B2']
    for ic, col_width in enumerate(format_cols):
        cell = ws.cell(row=1, column=ic+1)
        cell.alignment = l_alignment
        ws.column_dimensions[cell.column_letter].width = col_width
    return ws


def change_order_base_techno(new_columns_02):
    if 'База' in new_columns_02 and 'Техно' in new_columns_02:
        i_base = new_columns_02.index('База')
        i_techno = new_columns_02.index('Техно')
        # print(i_base, i_techno)
        if i_techno < i_base:
            new_columns_03 = [col for col in new_columns_02 if col not in ['Техно', 'База']]
            # print(new_columns_03)
            if i_base > 0: i_base -= 1
            new_columns_03.insert(i_base, 'Техно')
            new_columns_03.insert(i_base, 'База')


            return new_columns_03

        else: return new_columns_02
    else: return new_columns_02

def reorder_columns_by_models(new_columns_02, model_names):
    if model_names[0] in new_columns_02 and model_names[1] in new_columns_02:
        i_first = new_columns_02.index(model_names[0])
        i_second = new_columns_02.index(model_names[1])
        # print(i_base, i_techno)
        if i_second < i_first:
            new_columns_03 = [col for col in new_columns_02 if col not in model_names]
            # print(new_columns_03)
            if i_first > 0: i_first -= 1
            new_columns_03.insert(i_first, model_names[1])
            new_columns_03.insert(i_first, model_names[0])


            return new_columns_03

        else: return new_columns_02
    else: return new_columns_02

# def simplify_multi_index (df_p, tk_names, model_names):
#     '''
#     on enter pdDataFrame with columns
#     MultiIndex([('count',  'Техкарта БА КС база.xlsx'), ('count', 'Техкарта БА КС техно.xlsx')], names=[None, 'Файл Excel'])
#     '''
#     pp_lst = []
#     df_pp = df_p.reset_index()
#     for i_row, row in df_pp.iterrows():
#         pp_lst.append(row.values)
#     # print(pp_lst[:2])
#     cur_columns = list(df_pp.columns)
#     # cur_columns: [('Код раздела', ''), ('count', 'Техкарта БА КС база.xlsx'), ('count', 'Техкарта БА КС техно.xlsx')]
#     # print("cur_columns:", cur_columns)
#     new_columns = [v[0] if i_v in [0,3] else v[1] for i_v, v in enumerate(cur_columns)]
#     # print("new_columns:", new_columns)
#     cur_columns_02 = list(df_pp.columns[1:3])
#     # print("cur_columns_02:", cur_columns_02)
#     # new_columns_02 = ['База' if (str(col[1])==tk_names[0]) else 'Техно' for col in cur_columns_02]
#     new_columns_02 = [model_names[0] if (str(col[1])==tk_names[0]) else model_names[1] for col in cur_columns_02]
#     new_columns_02 = [new_columns[0]] + new_columns_02 # + [new_columns[-1]] #+ code_names_columns
#     # print("new_columns_02:", new_columns_02)

#     df_pp = pd.DataFrame(pp_lst, columns = new_columns_02)
#     # new_columns_03 = change_order_base_techno(new_columns_02)
#     new_columns_03 = reorder_columns_by_models(new_columns_02, model_names)
#     # print(f"new_columns_03: {new_columns_03}")
#     df_pp = df_pp[new_columns_03]

#     return df_pp



# def simplify_multi_index_02 (df_p, tk_names, model_names):
# # def simpl_multi_index_02 (df_p, tk_names, model_names):
#     '''
#     on enter pdDataFrame with columns
#     MultiIndex([('count',  'Техкарта БА КС база.xlsx'), ('count', 'Техкарта БА КС техно.xlsx')], names=[None, 'Файл Excel'])
#     '''
#     pp_lst = []
#     df_pp = df_p.reset_index()
#     for i_row, row in df_pp.iterrows():
#         pp_lst.append(row.values)
#     # print(pp_lst[:2])
#     cur_columns = list(df_pp.columns)
#     # print("cur_columns:", cur_columns)
#     new_columns = [v[0] if i_v in [0] else v[-1] for i_v, v in enumerate(cur_columns)]
#     # new_columns = [v[0]  for i_v, v in enumerate(cur_columns[:3])]
#     # print("new_columns:", new_columns)
#     # cur_columns_02 = list(df_pp.columns[1:3])
#     # print("cur_columns_02:", cur_columns_02)
#     # new_columns_02 = ['База' if (str(col[1])==tk_names[0]) else 'Техно' for col in cur_columns_02]
#     new_columns_02 = [model_names[0] if (str(col)==tk_names[0]) else model_names[1] for col in new_columns[1:]]
#     new_columns_02 = [new_columns[0]] + new_columns_02
#     # print("new_columns_02:", new_columns_02)

#     df_pp = pd.DataFrame(pp_lst, columns = new_columns_02)
#     # new_columns_03 = change_order_base_techno(new_columns_02)
#     new_columns_03 = reorder_columns_by_models(new_columns_02, model_names)
#     df_pp = df_pp[new_columns_03]

#     return df_pp


def extract_names_from_code_service(code, debug=False):
    section_name, type_name, class_name = None, None, None
    if (type(code)!= str) or ((type(code)==str) and (len(code)==0)): return section_name, type_name, class_name
    section_name = sevice_sections.get(code[0])
    if len(code)>=3:
        if code[0] == 'A':
            if debug: print('A')
            type_name = service_types_A.get(code[1:3])
        else: type_name = service_types_B.get(code[1:3])
        if len(code)>=6:
            if code[0] == 'A':
                class_name = service_classes_A.get(code[4:6])
            else: class_name = service_classes_B.get(code[4:7])
        else: return section_name, type_name, class_name
    else: return section_name, type_name, class_name

    return section_name, type_name, class_name


def services_analysis_02(
    df_services, tk_names, model_names, tk_code_name,
    path_tk_models_processed,
    analysis_subpart_code, analysis_subpart,
    indicator_col_name = 'Усредненная частота предоставления',
    agg_type = 'Среднее',

    ):

    codes_columns_services = ['Код раздела', 'Код типа', 'Код класса']
    code_names_columns_services = ['Раздел', 'Тип', 'Класс']
    services_mask_base = df_services['Файл Excel'] == tk_names[0]
    services_mask_techno = df_services['Файл Excel'] == tk_names[1]
    df_a = df_services[services_mask_base | services_mask_techno]
    # tk_name, model, analysis_part, analysis_part_code = 'Нейрохирургия',  'База', 'Услуги', 1
    # tk_name, model, analysis_part, analysis_part_code = tk_code_name,  'База', 'Услуги', 1
    # dictionaries_lst = [sevice_sections, (service_types_A, service_types_B), (service_classes_A, service_classes_B) ]
    diff_lst = []
    diff_df_services = []
    # code_names_columns_services = ['Раздел', 'Тип', 'Класс']
    n_bars_max_on_picture = 20
    # from matplotlib.colors import ListedColormap, BoundaryNorm
    colors=["#9b59b6", "#3498db", "#95a5a6", "#e74c3c", "#34495e", "#2ecc71"]
    cmap = ListedColormap(["#95a5a6", "#2ecc71"])

    for i_col, col_name in enumerate(codes_columns_services):
        diff_lst.append([])
        if agg_type == 'Среднее':
            df_p = df_a.groupby( ['Файл Excel', col_name, ] ).agg({indicator_col_name: ['mean']})\
                        .reset_index().pivot([col_name], ['Файл Excel'] ).fillna(0)
        elif agg_type == 'Сумма':
            df_p = df_a.groupby( ['Файл Excel', col_name, ] ).agg({indicator_col_name: ['sum']})\
                        .reset_index().pivot([col_name], ['Файл Excel'] ).fillna(0)
        # print(df_p.columns)
        # display(df_p.head(2))
        df_pp = simplify_multi_index_02 (df_p, tk_names, model_names)
        # df_pp = simpl_multi_index_02 (df_p, tk_names, model_names)
        # display(df_pp.head(2))
        kind = 'bar' #'kde' #'area' #'bar'
        title = '\n'.join([tk_code_name, 'Услуги', analysis_subpart]) #, indicator_col_name]) #, col_name])
        y_lim_min = 0

        print("!!! df_pp.shape[0]:", df_pp.shape[0])
        flag_pic_plotted = False
        if df_pp.shape[0] <= n_bars_max_on_picture:
            plt.figure(figsize=(25, 6), tight_layout=True)
            try:
                ax1 = df_pp.plot(kind= kind, x = col_name, rot=45, cmap = cmap)
                flag_pic_plotted = True
            except Exception as err:
                logger.error("!!! Ошибка данных")
                logger.error(f"{str(err)}")
                display(df_pp)
        else:
            plt.figure(figsize=(25, 10), tight_layout=True)
            try:
                max_v = max(df_pp[model_names[0]].max(), df_pp[model_names[1]].max())
                min_v = min(df_pp[model_names[0]].min(), df_pp[model_names[1]].min())
                try:
                    delta_v = (max_v - min_v)/10
                    for i_max in range(10):
                        # df_pp1 = df_pp[(df_pp['База']>=y_lim_min + i_max) | (df_pp['Техно']>=y_lim_min + i_max)]

                        df_pp1 = df_pp[(df_pp[model_names[0]]>=y_lim_min + i_max*delta_v) | (df_pp[model_names[1]]>=y_lim_min + i_max*delta_v)]

                        if df_pp1.shape[0] <= n_bars_max_on_picture:
                            print(f"i_max: {i_max}, df_pp1.shape[0]: {df_pp1.shape[0]}")
                            ax1 = df_pp1.plot(kind= kind, x = col_name, rot=45, cmap = cmap) #, y_lim= (y_lim_min + i_max,100))
                            flag_pic_plotted = True
                            break
                except Exception as err:
                    logger.error("!!! Ошибка данных")
                    logger.error(f"{str(err)}")
                    display(df_pp)
            except Exception as err:
                print(str(err))
                try:
                    ax1 = df_pp.plot(kind= kind, x = col_name, rot=45, cmap = cmap)
                except Exception as err:
                    logger.error("!!! Ошибка данных")
                    logger.error(f"{str(err)}")
                    display(df_pp)
        legend_list = model_names
        if flag_pic_plotted:
            ax1.legend(legend_list, loc='best',fontsize=8)
            plt.title(title, fontsize=8)
            plt.xticks(fontsize=8)
            plt.yticks(fontsize=8)
            plt.xlabel(col_name, fontsize=8)
            # plt.ylabel('Количество', fontsize=8)
            plt.ylabel(agg_type, fontsize=8)

            # fn_img = f"{analysis_part_code:02d}_{analysis_part}_{i_col:02d}.jpg"
            fn_img = f"01_Услуги_{analysis_subpart_code:02d}_{analysis_subpart}_{i_col:02d}.jpg" #.replace(' ','_')

            # plt.savefig(os.path.join(path_tk_models_processed, tk_code_name, fn_img), bbox_inches='tight')
            plt.savefig(os.path.join(path_tk_models_processed, tk_code_name, 'img', fn_img), bbox_inches='tight')
            # plt.savefig(path_tk_models_processed + tk_code_name + '/' + fn_img, bbox_inches='tight')
            plt.show()
        try:
            diff_df_services.append(def_differencies(
                                 df_pp, tk_names, model_names,
                                 code_names_columns = code_names_columns_services,
                                 function_extract_names = extract_names_from_code_service))
            display(diff_df_services[i_col])
        except Exception as err:
            diff_df_services.append(None)
            logger.error(str(err))
            logger.error(f"Данные анализа об отличиях не выводятся из-за некорректных входных данных")
    return diff_df_services

def extract_name_groups_ATH(s, debug = False):
    ath_anatomy_code, ath_anatomy, ath_therapy_code, ath_therapy, ath_pharm_code, ath_pharm, ath_chemical_code, ath_chemical = \
        None, None, None, None, None, None, None, None
    if type(s) is None or ((type(s)==float) and np.isnan(s)) or (type(s)!=str)  or ((type(s)==str) and (len(s)==0)):
        return None, None, None, None, None, None, None, None
    ath_anatomy_code = s[0]
    ath_anatomy = dict_ath_anatomy.get(ath_anatomy_code)
    if len(s)>=3:
        ath_therapy_code = s[0:3]
        ath_therapy = dict_ath_therapy.get(ath_therapy_code)
        if len(s)>=4:
            ath_pharm_code = s[0:4]
            ath_pharm = dict_ath_pharm.get(ath_pharm_code)
            if len(s)>=5:
                ath_chemical_code = s[0:5]
                ath_chemical = dict_ath_chemical.get(ath_chemical_code)
    return ath_anatomy, ath_therapy, ath_pharm, ath_chemical


# def update_excel_by_analysis_02_options(
#     diff_df_services_02, diff_LP_df_02,
#     path_tk_models_processed, tk_code_name, fn_TK_save,
#     cmp_sections
#     ):

#     wb = load_workbook(os.path.join(path_tk_models_processed, tk_code_name, fn_TK_save))
#     # tk_name, model, analysis_part, analysis_part_code = 'Нейрохирургия',  'База', 'Услуги', 1
#     if len(diff_df_services_02)==0: diff_df_services_02 = None
#     if len(diff_LP_df_02)==0: diff_LP_df_02 = None

#     df_diff = [diff_df_services_02, diff_LP_df_02, None]
#     cols_width_analysis = [[10,7,7,7,30,30,30], [10,7,7,7,30,30,30,30], None]
#     interval_row = 1
#     analysis_subpart_lst = [ [(2, 'Частота'), (3, 'Кратность'), (4, 'УЕТ 1'), (5, 'УЕТ 2')],
#                              [(2, 'Частота'), (3, 'Кратность'), (4, 'Количество')]
#     ]
#     for i_p, analysis_part in enumerate(['Услуги', 'ЛП']): #, 'РМ']):
#         if analysis_part in cmp_sections:
#             for i_sp, (analysis_subpart_code, analysis_subpart) in enumerate(analysis_subpart_lst[i_p]):
#                 fn_img_lst = glob.glob(os.path.join(
#                     path_tk_models_processed, tk_code_name, 'img') + f"/{i_p+1:02d}_{analysis_part}_{analysis_subpart_code:02d}_{analysis_subpart}*.jpg")
#                     # path_tk_models_processed, tk_code_name) + f"/{i_p+1:02d}_{analysis_part}_{analysis_subpart_code:02d}_{analysis_subpart}*.jpg")
#                     # path_tk_models_processed, tk_code_name) + f"/{i_p+1:02d}_{analysis_part}_{i_sp:02d}_{analysis_subpart.replace(' ', '_')}*.jpg")
#                 print("fn_img_lst:", len(fn_img_lst), fn_img_lst)
#                 sheet_name = analysis_part + '_Анализ_' + analysis_subpart #.replace(' ', '_')
#                 sheet_names = wb.get_sheet_names()
#                 if sheet_name in sheet_names:
#                     # wb.remove_sheet(sheet_name)
#                     wb.remove(wb[sheet_name])
#                 wb.create_sheet(sheet_name)
#                 ws = wb[sheet_name]
#                 if cols_width_analysis[i_p] is not None:
#                     ws = format_excel_cols_short(ws, cols_width_analysis[i_p], auto_filter=False)
#                 # cell = ws['A1']
#                 # font_size = cell.font.sz
#                 cell_height = 20 # опытным путем
#                 cell_height = 17 # опытным путем

#                 images_total_height = 0
#                 images_total_rows = 0
#                 explain_rows = 0
#                 interval_rows = 0

#                 for i_f, fn_img in enumerate(fn_img_lst):
#                     img = drawing.image.Image(fn_img)
#                     anchor = f"A{images_total_rows + explain_rows+1}"
#                     ws.add_image(img, anchor)
#                     # img_rows = int(img.height//cell_height   + 1) # + interval_row
#                     img_rows = img.height//cell_height   + 1 + 2*interval_row
#                     images_total_rows += img_rows
#                     for _ in range(img_rows): ws.append([None])

#                     if df_diff[i_p] is not None:
#                         # cell = ws[anchor]
#                         # print(f"i_p: {i_p}, len(df_diff[i_p]):", len(df_diff[i_p]))
#                         if df_diff[i_p][i_sp] is not None:
#                         # if df_diff[i_p][i_f] is not None:
#                             try:
#                                 # ws.append(list(df_diff[i_p][i_f].columns))
#                                 # for i_row, row in df_diff[i_p][i_f].iterrows():
#                                 ws.append(list(df_diff[i_p][i_sp][i_f].columns))
#                                 for i_row, row in df_diff[i_p][i_sp][i_f].iterrows():
#                                     ws.append(list(row.values))
#                                 explain_rows += df_diff[i_p][i_sp][i_f].shape[0] + 1 + 2*interval_row
#                             except Exception as err:
#                                 print(err)
#                                 # print(type(df_diff[i_p][i_f]), df_diff[i_p][i_f])
#                         else:
#                             for i_row, row in range(2*interval_row):
#                                 ws.append([None])
#                             explain_rows += 2*interval_row

#             # print(img.height, img_rows, images_total_rows, explain_rows)

#     wb.save(os.path.join(path_tk_models_processed, tk_code_name, fn_TK_save))
#     logger.info(f"Файл '{fn_TK_save}' дополнен данными анализа и сохранен в '{os.path.join(path_tk_models_processed, tk_code_name)}'")




# import pandas as pd
from utils_form_tk_enrichment_options import preprocess_tkbd_options
from utils_io import form_str_date, format_excel_sheet_cols, add_sheet_to_excel_from_df

def group_services(df_services, freq_threshold, group_code_col ='Код типа', group_name_col='Тип'):

    code_col = 'Код услуги по Номенклатуре медицинских услуг (Приказ МЗ № 804н)'
    name_col = 'Наименование услуги по Номенклатуре медицинских услуг (Приказ МЗ №804н)'
    freq_col = 'Усредненная частота предоставления'
    multi_col = 'Усредненная кратность применения'
    head_cols = ['Профиль', 'Код ТК', 'Наименование ТК', 'Модель пациента', 'Файл Excel']
    main_cols = ['Код услуги по Номенклатуре медицинских услуг (Приказ МЗ № 804н)',
       'Наименование услуги по Номенклатуре медицинских услуг (Приказ МЗ №804н)',
       'Усредненная частота предоставления',
       'Усредненная кратность применения'] #, 'Код раздела', 'Раздел', 'Код типа'

    df_g_services1 = df_services[df_services[freq_col]>=freq_threshold]
    df_g_services2 = df_services[df_services[freq_col]<freq_threshold]
    print(df_services.shape, df_g_services1.shape, df_g_services2.shape)
    # display(df_g_services2.groupby(group_col).sum(freq_col))
    # display(df_g_services2[list(df_g_services1.columns)].groupby(group_col).sum(freq_col).reset_index().head())
    groupby_cols = head_cols + [group_code_col, group_name_col]
    df_g_services2_g = df_g_services2[ groupby_cols + [freq_col, multi_col]
                                      ].groupby(groupby_cols).agg({freq_col: 'sum', multi_col: 'mean'}).reset_index()
    # print("df_g_services2_g.columns:", df_g_services2_g.columns)
    # display(df_g_services2_g.head(2))
    df_g_services2_g[code_col] = df_g_services2_g[group_code_col].apply(lambda x: x + ('.AA' if x.startswith('A') else '.BBB'))
    df_g_services2_g[name_col] = df_g_services2_g[group_name_col].apply(lambda x: x.capitalize() + '.*')

    # print("df_g_services1.columns:", df_g_services1.columns)
    # print(df_g_services2_g.shape)
    # display(df_g_services2_g.head(2))
    df_g_services2_g.drop(columns= [group_code_col, group_name_col], inplace=True)
    df_g_services = pd.concat([df_g_services1[list(df_g_services2_g.columns)], df_g_services2_g])

    df_g_services = df_g_services[head_cols + [code_col, name_col, freq_col, multi_col]]
    df_g_services = df_g_services.sort_values(by=[code_col], ascending=True)
    print(df_g_services.shape)

    return df_g_services

def group_items(df_services, df_LP, df_RM,
        freq_threshold,
):
    df_g_services, df_g_LP, df_g_RM = None, None, None
    if df_services is not None:
        df_g_services = group_services(df_services, freq_threshold)
    return df_g_services, df_g_LP, df_g_RM
def form_tk_options( data_source_dir, data_processed_dir, supp_dict_dir,
                        fn_check_file1, cmp_cols_file_01,
                        fn_smnn_pickle, cmp_sections,
                        profile='profile_test', tk_code=7777777, tk_name='tk_test',
                        models = ['Факт', ],
                        freq_threshold=0.05,

                        ):

    if fn_check_file1 is None:
        logger.error(f"Выберите название файла: в параметрах запуска программы")
        sys.exit(2)
    # df_services, df_LP, df_RM = preprocess_tkbd_options(data_source_dir, fn_tk_bd, data_processed_dir, supp_dict_dir, fn_smnn_pickle)
    df_services, df_LP, df_RM = preprocess_tkbd_options(
                data_source_dir, data_processed_dir, supp_dict_dir,
                fn_check_file1,
                cmp_cols_file_01,

                fn_smnn_pickle, cmp_sections,
                profile, tk_code, tk_name,
                models,
                save_enriched=False,
                )
    if df_services is not None: display(df_services.head(2))
    if df_LP is not None: display(df_LP.head(2))
    if df_RM is not None: display(df_RM.head(2))

    df_g_services, df_g_LP, df_g_RM = group_items(df_services, df_LP, df_RM, freq_threshold)
    df_g_lst, sheet_names_lst, sheet_names_g_lst = [], [], []
    if df_g_services is not None:
        df_g_lst.append(df_g_services)
        sheet_names_lst.append(f"Услуги")
        sheet_names_g_lst.append(f"Услуги_грп_{str(freq_threshold)}")
    if df_g_LP is not None:
        df_g_lst.append(df_g_LP)
        sheet_names_lst.append(f"ЛП")
        sheet_names_g_lst.append(f"ЛП_грп_{str(freq_threshold)}")
    if df_g_RM is not None:
        df_g_lst.append(df_g_RM)
        sheet_names_lst.append(f"РМ")
        sheet_names_g_lst.append(f"РМ_грп_{str(freq_threshold)}")
    str_date = form_str_date ()
    fn_tk = 'groupped_tk_' + str_date + '.xlsx'
    fn_tk = f"groupped_tk_{profile}_{tk_code}.xlsx"
    fn_save = save_df_lst_to_excel(df_g_lst, sheet_names_g_lst, data_processed_dir, fn_tk)
    sections = ['Услуги', 'ЛП', 'РМ']
    format_cols = [
        [30,15,30, 30,30, 15,30, 20,20],
        [],
        [],
    ]
    for i_sh, sheet_name in enumerate(sheet_names_g_lst):
        if i_sh==0:
            add_sheet_to_excel_from_df(df_g_lst[i_sh], sheet_name,  data_source_dir, fn_check_file1, data_processed_dir, fn_save, index=False)
        else:
            add_sheet_to_excel_from_df(df_g_lst[i_sh], sheet_name,  data_processed_dir, fn_save, data_processed_dir, fn_save, index=False)
        format_excel_sheet_cols(data_processed_dir, fn_save, format_cols[sections.index(sheet_names_lst[i_sh])], sheet_name=sheet_name)

    logger.info(f"Файл '{fn_save}' сохранен в директорию '{data_processed_dir}'")
    return df_g_services, df_g_LP, df_g_RM
