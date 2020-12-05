import pandas as pd
import openpyxl
import module.forms_maker._0420503_Prirost as prst
global log

# ===========================================================================
def main(id_fond, path_to_report, file_Avancore_prirost, file_xbrl, report):
    # Загружаем данные
    # ---------------------------------------------------------
    # Загружаем данные из файла Аванкор
    # (файл содержит только одну вкуладку: 'TDSheet')
    df_avancor = pd.read_excel(path_to_report + '/' + file_Avancore_prirost,
                               header=None)
    # устанавливаем начальный индекс не c '0', а c '1'
    df_avancor.index += 1
    df_avancor.columns += 1
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    wb_xbrl = openpyxl.load_workbook(filename=(path_to_report + '/' + file_xbrl))
    # ---------------------------------------------------------
    # Копируем данные в формы Прироста
    prst.scha_prirost(id_fond, wb_xbrl, df_avancor)
    # ---------------------------------------------------------
    # Сохраняем результаты в файл отчетности xbrl
    try:
        wb_xbrl.save(path_to_report + '/' + file_xbrl)
        print('-------------------ГОТОВО!!!----------------------')

    except PermissionError:
        log.error(f'{"=" * 100}\n'
                  f'ОШИБКА ДОСТУПА К ФАЙЛУ\n'
                  f'(файл открыт в другой программе - закройте файл!)\n'
                  f'{"=" * 100} ')


# ============================================================================

if __name__ == '__main__':
    pass
