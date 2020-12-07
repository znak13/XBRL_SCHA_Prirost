"""Загружаем данные из файлов """

import openpyxl
from module.globals import *


def load_pif_info(file_name=pif_info,
                  path_2file=dir_shablon):
    """Загружаем данные из файла с информацией о фондах"""
    file_name = path_2file + file_name
    wb = openpyxl.load_workbook(file_name, data_only=True)
    # data_only=True - загружаем данные (без формул)

    return wb


# ==========================================================================
if __name__ == "__main__":
    pass

