import pandas as pd
import shutil
import openpyxl
import builtins

from module.globals import *
from module import logger

import module.forms_maker._0420502_SCHA as scha
import module.forms_maker._0420502_Rasshifr as rf
import module.forms_maker._0420502_Podpisant as pp
import module.forms_maker._0420502_Zapiski as zap
import module.forms_maker._0420503_Prirost as prst


# ================================================================
def main(id_fond, path_to_report, file_fond_name, file_Avancore_scha):
    # ----------------------------------------------------------
    # Включаем логировние
    # log = logger.create_log(path=path_to_report + '/',
    #                         file_log=file_fond_name + '_' + id_fond + log_endName_scha,
    #                         file_debug=file_fond_name + '_' + id_fond + debug_endName_scha)
    # # устанавливаем 'log' как глобальную переменную (включая модули)
    # builtins.log = log
    # ----------------------------------------------------------
    # Загрузка файла с Идентификаторами
    df_identifier = pd.read_excel(dir_shablon + fileID,
                                  sheet_name=None,
                                  index_col=None,
                                  header=None)
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
    shutil.copyfile(dir_shablon + fileShablon, path_to_report + '/' + file_fond_name)
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=(path_to_report + '/' + file_fond_name))
    # ----------------------------------------------------------
    # Формимируем файл-xbrl-СЧА:
    # Формируем итоговые формы СЧА
    scha.scha(wb, id_fond, df_identifier, df_avancor)
    # Формируем итоговые формы-расшифровки
    rf.rashifr(wb, df_avancor, df_identifier, id_fond)
    # Формирование форм СЧА - падписанты
    pp.podpisant(wb, df_avancor, df_identifier, id_fond)
    # Формирование форм СЧА - пояснительные записки
    zap.zapiski_new(wb, df_avancor, id_fond)
    # Удаляем формы-Прироста, которые не заполняются
    # (остальные формы Прироста формируются отдельно)
    prst.prirost(wb)
    # Сохраняем результат
    wb.save(path_to_report + '/' + file_fond_name)


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == '__main__':
    pass
