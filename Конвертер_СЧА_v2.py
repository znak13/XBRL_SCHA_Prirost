import pandas as pd
import shutil
import openpyxl

from module.globals import *
import module.forms_maker._0420502_SCHA as scha
import module.forms_maker._0420502_Rasshifr as rf
import module.forms_maker._0420502_Podpisant as pp
import module.forms_maker._0420502_Zapiski as zap
import module.forms_maker._0420503_Prirost as prst

from module.dataCheck import checkSheetsInFileID


# ================================================================
def main(id_fond, path_to_report, file_new_name, file_Avancore_scha):
    # ----------------------------------------------------------
    # Загрузка файла с Идентификаторами
    df_identifier = pd.read_excel(dir_shablon + fileID,
                                  sheet_name=None,
                                  index_col=None,
                                  header=None)
    # Проверка файла с идентификаторами на предмет наличия всех вкладок и лишних вкладок
    # (111 - pas. - защита структуры в файле с идентификаторами)
    checkSheetsInFileID(df_identifier)
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


    # ----------------------------------------------------------
    # Формимируем файл-xbrl-СЧА:
    # Формируем итоговые формы СЧА
    scha.scha(wb, id_fond, df_identifier, df_avancor)
    # Формируем итоговые формы-расшифровки
    rf.rashifr(wb, df_avancor, df_identifier, id_fond)
    # Формирование форм СЧА - падписанты
    pp.podpisant(wb, df_avancor, id_fond)
    # Формирование форм СЧА - пояснительные записки
    zap.zapiski_new(wb, df_avancor, id_fond)
    # Удаляем формы-Прироста, которые не заполняются
    # (остальные формы Прироста формируются отдельно)
    prst.prirost(wb)
    # Сохраняем результат
    wb.save(path_to_report + '/' + file_new_name)


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
if __name__ == '__main__':
    pass
