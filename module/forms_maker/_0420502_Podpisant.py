""" Формирование форм СЧА - падписанты"""
import pandas as pd

import module.functions as fun
import module.adjustments as adj
import module.dataCheck as dCheck

from module.globals import *
global log
# Формы:
# 0420502 Справка о стоимости _56	SR_0420502_Podpisant
# 0420502 Справка о стоимости _57	SR_0420502_Podpisant_spec_dep

# def urlForms():
#     return ['SR_0420502_Podpisant',  # 0
#             'SR_0420502_Podpisant_spec_dep']  # 1
#
#
# def podpisant(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS, errors_file):
#     urlSheets = fun.codesSheets(wb)
#     # ================================================================
#     # Корректируем формы в зависимости от их особенностей:
#     # ================================================================
#     """0420502 Справка о стоимости _56	SR_0420502_Podpisant"""
#     shortURL = urlForms()[0]
#     sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)
#     ws = wb[sheetName]
#
#     # Записываем ФИО подписантов
#     adj.corrector_Podpisant_1_(ws, df_matrica, df_avancor, shortURL, ERRORS, errors_file)
#
#     # Заменяем на полное ФИО подписантов
#     cell_with_fio = 'B7'
#     adj.corrector_Podpisant_2_(ws, df_identifier, sheetName, cell_with_fio, ERRORS, errors_file)
#
#     # ================================================================
#     """0420502 Справка о стоимости _57	SR_0420502_Podpisant_spec_dep"""
#     shortURL = urlForms()[1]
#     sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)
#     ws = wb[sheetName]
#
#     # Записываем ФИО подписантов
#     adj.corrector_Podpisant_1_(ws, df_matrica, df_avancor, shortURL, ERRORS, errors_file)
#     # Заменяем на полное ФИО подписантов
#     cell_with_fio = 'B8'
#     adj.corrector_Podpisant_2_(ws, df_identifier, sheetName, cell_with_fio, ERRORS, errors_file)
#     # Проставляем id-Фонда
#     ws.cell(8, 1).value = id_fond
#     # Проставляем реквизиты СпецДепа
#     adj.corrector_Podpisant_3_(ws, df_identifier, id_fond)
#
#     # ================================================================


def podpisant(wb, df_avancor, id_fond):

    # **********************************************************************************
    def fioShort(ws, AvancoreTitle, avancor_fio_col, cell_fio):
        """Копируем короткое-ФИО подписанта"""
        # Формы:
        # 0420502 Справка о стоимости _56   SR_0420502_Podpisant
        # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep

        # количество строк в таблице Аванкор
        index_max = df_avancor.shape[0]
        # Номер строки с названием раздела в файле Аванкор
        avancor_title_row = fun.razdel_name_row(df_avancor, AvancoreTitle, index_max)

        # ФИО подписанта
        fio = df_avancor.loc[avancor_title_row, avancor_fio_col]

        # записываем в форму xbrl
        row_fio, col_fio = fun.coordinate(cell_fio)
        ws.cell(row_fio, col_fio).value = fio

    # **********************************************************************************
    def fioFull(ws, cell_with_fio):
        """ Вставляем в ячейки с подписантами полное ФИО вместо сокращенного ФИО"""

        # Формы:
        # 0420502 Справка о стоимости _56   SR_0420502_Podpisant
        # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep

        def insert_fio_full():
            """ Вставляем полное ФИО в ячейку"""

            def searche_fio_full():
                """ Ищем полное ФИО в файле с Идентификаторами """

                fio_full = None
                # Удаляем лишние пробелы из ФИО (если пробелы есть)
                try:
                    fio_short_bp = fio_short.replace(' ', '')
                except AttributeError:
                    log.error(f'"{ws.title}" --> ФИО подписанта: "{fio_short}" '
                              f'- не заполнено')
                    return fio_full
                except:
                    log.error(f'"{ws.title}" - Неизвестная ошибка')
                    return fio_full

                # Удаляем лишние пробелы из ФИО (если пробелы есть)
                for n in range(len(ws_fio['ФИО_кратко'])):
                    # n += 1
                    ws_fio['ФИО_кратко'][n] = ws_fio['ФИО_кратко'][n].replace(' ', '')
                try:  # находим нужное ФИО
                    fio = ws_fio[ws_fio['ФИО_кратко'] == fio_short_bp]
                    # переиндексируем df, начиная с '0'
                    fio.index = range(0, len(fio))
                    fio_full = fio.loc[0, 'ФИО_полностью']
                except KeyError:
                    log.error(f'"{ws.title}" --> полное ФИО подписанта: "{fio_short}" '
                              f'- не найдено')
                except:
                    log.error(f'"{ws.title}" --> Неизвестная ошибка')
                return fio_full

            # ..........................................................
            fio_short = ws[cell_with_fio].value
            fio_full = searche_fio_full()
            # Если находим полное ФИО, то вставляем его в форму xbrl
            if fio_full:
                ws[cell_with_fio].value = fio_full
                # print(fio_full)
        # --------------------------------------------------------------
        # # Загружаем нужную страницу файла с Идентификаторами
        # ws_fio = df_identifier['ФИО']
        # # Выбираем в качестве заголовкой столбцов первую строку
        # ws_fio.columns = ws_fio.iloc[0]
        # # удаляем нулевую строку - она дублирует заголовки
        # ws_fio = ws_fio.drop(ws_fio.index[[0]])
        # # переиндексируем строки, начиная с "0"
        # ws_fio.index -= 1

        ws_fio = pd.read_excel(dir_shablon + pif_info,
                               sheet_name=pif_info_sheet_fio,
                               index_col=False,
                               header=0)
        insert_fio_full()

    # **********************************************************************************
    def podpisant_56():
        """0420502 Справка о стоимости _56	SR_0420502_Podpisant"""

        shortURL = 'SR_0420502_Podpisant'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        print(f'{sheetName} - {shortURL}')
        cell_fio = 'B7'

        AvancoreTitle = 'Руководитель      акционерного      инвестиционного\n' \
                        'фонда (управляющей компании паевого инвестиционного\n' \
                        'фонда)  (лицо, исполняющее обязанности руководителя\n' \
                        'акционерного инвестиционного фонда\n' \
                        '(управляющей компании паевого инвестиционного фонда)'
        # Номер колонки ячейки с ФИО в таблице Аванкор
        avancor_fio_col = 'J'
        avancor_fio_col = fun.column_index_from_string(avancor_fio_col)

        # Записываем короткое-ФИО подписанта
        fioShort(ws, AvancoreTitle, avancor_fio_col, cell_fio)
        # Заменяем короткое-ФИО на полное-ФИО подписанта
        fioFull(ws, cell_fio)

    # **********************************************************************************
    def podpisant_57():
        """0420502 Справка о стоимости _57	SR_0420502_Podpisant_spec_dep"""

        shortURL = 'SR_0420502_Podpisant_spec_dep'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        print(f'{sheetName} - {shortURL}')
        cell_fio = 'B8'

        AvancoreTitle = 'Уполномоченное лицо специализированного депозитария\n' \
                        'акционерного инвестиционного фонда (паевого инвестиционного фонда)'
        # Номер колонки ячейки с ФИО в таблице Аванкор
        avancor_fio_col = 'J'
        avancor_fio_col = fun.column_index_from_string(avancor_fio_col)

        # Записываем короткое-ФИО подписанта
        fioShort(ws, AvancoreTitle, avancor_fio_col, cell_fio)
        # Заменяем короткое-ФИО на полное-ФИО подписанта
        fioFull(ws, cell_fio)

        # Проставляем id-Фонда
        ws['A8'].value = id_fond
        # Проставляем реквизиты СпецДепа
        # adj.corrector_Podpisant_3_(ws, df_identifier, id_fond)
        adj.corrector_Podpisant_3_v2(ws, id_fond)


    # **********************************************************************************
    urlSheets = fun.codesSheets(wb)  # словарь - "код вкладки":"имя вкладки"
    podpisant_56()
    podpisant_57()


# %%

if __name__ == "__main__":
    pass
