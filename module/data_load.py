"""Загружаем данные из файлов 'Аванкор' и 'Идентификаторы',
а также создаем новый файл отчетности xbrl
"""

import pandas as pd
import openpyxl
import shutil
from module.globals import *


def load_info_from_files(dir_shablon,
                         fileID,
                         path_to_report,
                         file_Avancore_scha,
                         file_new_name):
    # ----------------------------------------------------------
    # Загрузка файла с Идентификаторами
    df_identifier = pd.read_excel(dir_shablon + fileID,
                                  sheet_name=None,
                                  index_col=None,
                                  header=None)
    # Проверка файла с идентификаторами на предмет наличия всех вкладок и лишних вкладок
    # (111 - pas. - защита структуры в файле с идентификаторами)
    # checkSheetsInFileID(df_identifier)
    # ----------------------------------------------------------
    # Загрузка файла-Аванкор-СЧА
    df_avancor = pd.read_excel(path_to_report + '/' + file_Avancore_scha,
                               index_col=None,
                               header=None)
    # устанавливаем начальный индекс не c 0, а c 1 (так удобнее)
    df_avancor.index += 1
    df_avancor.columns += 1
    # ----------------------------------------------------------
    # Создаем новый файл отчетности 'file_fond_name',
    # создав копию шаблона 'file_shablon'
    shutil.copyfile(dir_shablon + fileShablon, path_to_report + '/' + file_new_name)
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=(path_to_report + '/' + file_new_name))

    return df_identifier, df_avancor, wb


def load_pif_info(file_name=pif_info,
                  path_2file=dir_shablon):
    """Загружаем данные из файла с информацией о фондах"""
    file_name = path_2file + file_name
    wb = openpyxl.load_workbook(file_name)

    return wb


# ==========================================================================
if __name__ == "__main__":
    pass

    import os
    os.chdir('D:/Clouds/YandexDisk/Git/XBRL_SCHA_Prirost')

    wb = load_pif_info(file_name=pif_info,
                  path_2file=dir_shablon)
    ws = wb['ФИО']
