""" Корректировки содержания форм отчетности после копирования данных из файла Аванкор"""

from openpyxl.styles import Alignment
from module.functions import razdel_name_row, coordinate
from module.functions import codesSheets, listSheetsName, sheetNameFromUrl
from module.analiz_data import analiz_data_data, toFixed

global log

def corrector_scha_01(wb, df_id, id_fond, shortURL=None):
    """ Меняем местами значения ячеек и добавляем id фонда и УК"""
    # Форма:
    # 0420502 Справка о стоимости чис   SR_0420502_R1

    sheetName = sheetNameFromUrl(codesSheets(wb), shortURL)
    ws_change = wb[sheetName]

    # Меняем значения ячеек, т.к. в таблице xbrl ячейки расположены
    # в ином порядке, чем в файле, созданном Аванкор
    ws_change.cell(10, 5).value, ws_change.cell(10, 6).value = \
        ws_change.cell(10, 6).value, ws_change.cell(10, 5).value
    ws_change.cell(10, 5).value, ws_change.cell(10, 7).value = \
        ws_change.cell(10, 7).value, ws_change.cell(10, 5).value

    # Записываем в форму идентификаторы: фонда и УК
    ws_change.cell(10, 1).value = id_fond
    ws_change.cell(10, 2).value = df_id['УК АИФ ПИФ'].loc[1, 0]


# %%
def corrector_scha_03to10(wb, shortURL=None):
    """ Проставляем нулевые значения в формах: ч_3 - ч_10 """

    # Формы:
    # 0420502 Справка о стоимости ч_3    SR_0420502_R3_P1
    # 0420502 Справка о стоимости ч_4    SR_0420502_R3_P4
    # 0420502 Справка о стоимости ч_5    SR_0420502_R3_P2
    # 0420502 Справка о стоимости ч_6    SR_0420502_R3_P3
    # 0420502 Справка о стоимости ч_7    SR_0420502_R3_P5
    # 0420502 Справка о стоимости ч_8    SR_0420502_R3_P6
    # 0420502 Справка о стоимости ч_9    SR_0420502_R3_P7
    # 0420502 Справка о стоимости _10    SR_0420502_R3_P8

    def insert_00(ws):
        """ Проставляем нулевые значения в форме """

        row_begin = 10
        col_1 = 3  # номер первой колонки с данными

        for row in range(row_begin, ws.max_row + 1):
            cell_1 = ws.cell(row, col_1).value
            cell_2 = ws.cell(row, col_1 + 1).value
            cell_3 = ws.cell(row, col_1 + 2).value
            cell_4 = ws.cell(row, col_1 + 3).value

            # Форматируем ячейки
            for i in range(4):
                ws.cell(row, col_1 + i).alignment = Alignment(horizontal='right')

            if cell_1 or cell_2:  # Если есть данные в одной из первых ячеек
                if not cell_1:  # если нет в первой
                    ws.cell(row, col_1).value = '0.00'
                if not cell_2:  # если нет во второй
                    ws.cell(row, col_1 + 1).value = '0.00'
                if not cell_3:  # если нет в третьей
                    ws.cell(row, col_1 + 2).value = '0.00'
                if not cell_4:  # если нет в червертой
                    ws.cell(row, col_1 + 3).value = '0.00'

    # проставляем нули
    # if not newCode:
    #     shortURLs = ['SR_0420502_R3_P1',
    #                  'SR_0420502_R3_P4',
    #                  'SR_0420502_R3_P2',
    #                  'SR_0420502_R3_P3',
    #                  'SR_0420502_R3_P5',
    #                  'SR_0420502_R3_P6',
    #                  'SR_0420502_R3_P7',
    #                  'SR_0420502_R3_P8']
    #     sheets = listSheetsName(wb, shortURLs)
    #
    #     for sheet in sheets:
    #         ws = wb[sheet]
    #         insert_00(ws)
    # else:

    sheet = sheetNameFromUrl(codesSheets(wb), shortURL)
    ws = wb[sheet]
    insert_00(ws)


# %%
def corrector_scha_03to12_(wb_xbrl, df_avancor, cell_period, shortURL=None):
    """ Вставляем данные о периоде отчетности """
    # Формы:
    # 0420502 Справка о стоимости ч_3   SR_0420502_R3_P1
    # 0420502 Справка о стоимости ч_4   SR_0420502_R3_P4
    # 0420502 Справка о стоимости ч_5   SR_0420502_R3_P2
    # 0420502 Справка о стоимости ч_6   SR_0420502_R3_P3
    # 0420502 Справка о стоимости ч_7   SR_0420502_R3_P5
    # 0420502 Справка о стоимости ч_8   SR_0420502_R3_P6
    # 0420502 Справка о стоимости ч_9   SR_0420502_R3_P7
    # 0420502 Справка о стоимости _10   SR_0420502_R3_P8
    # 0420502 Справка о стоимости _11   SR_0420502_R3_P9
    # 0420502 Справка о стоимости _12   SR_0420502_R4

    period_begin = analiz_data_data(df_avancor.loc[18, 1])
    period_end = analiz_data_data(df_avancor.loc[18, 5])
    sheet_xbrl_period = sheetNameFromUrl(codesSheets(wb_xbrl), shortURL)

    ws_period = wb_xbrl[sheet_xbrl_period]
    # cell_period = df_matrica.loc[shortURL, 'cell_period']
    ws_period[cell_period].value = period_begin + ', ' + period_end


# %%
def corrector_scha_13_(id_fond, ws, df_avancor, df_id):
    """ Копируем количество паев с нужной точностью """
    # форма:
    # 0420502 Справка о стоимости _13

    # Точность указания кол-ва паев
    df_id_ws = df_id['ПИФ']
    fix = df_id_ws[df_id_ws[0] == id_fond][2].values[0]

    today_row, today_col = coordinate('I144')
    previous_row, previous_col = coordinate('K144')

    today = toFixed(df_avancor.loc[today_row, today_col], digits=fix)
    previous = toFixed(df_avancor.loc[previous_row, previous_col], digits=fix)

    ws.cell(10, 3).value = today
    ws.cell(10, 4).value = previous

    # Форматируем ячейки
    for row in range(9, ws.max_row + 1):
        for col in range(3, ws.max_column + 1):
            ws.cell(row, col).alignment = Alignment(horizontal='right')


# %%
# def corrector_scha_01_14_15_53(wb):
#     """ Убираем лишние нули (".00"), которые могут появлиться в результате анализа данных """
#
#     # формы:
#     # 0420502 Справка о стоимости чис   SR_0420502_R1
#     # 0420502 Справка о стоимости _14   SR_0420502_Rasshifr_Akt_P1_P1
#     # 0420502 Справка о стоимости _15   SR_0420502_Rasshifr_Akt_P1_P2
#     # 0420502 Справка о стоимости _53   SR_0420502_Rasshifr_Akt_P7_3
#
#     def correct(row, col):
#         # разбиваем на две части и оставляем только первую часть (без нулей)
#         ws.cell(row, col).value = str(ws.cell(row, col).value).split('.')[0]
#
#     urlSheet = codesSheets(wb)
#     # Корректируем реквизиты фонда: "№ лицензии", "ОКПО"
#     # названия вкладок
#     sheet_rekvizits = sheetNameFromUrl(urlSheet, 'SR_0420502_R1')
#     ws = wb[sheet_rekvizits]
#     correct(10, 4)  # "№ лицензии"
#     correct(10, 5)  # "ОКПО"
#
#     # Корректируем номер кредитной организации (поле:"Регистрационный номер кредитной организации")
#     # названия вкладок
#     shortURLs = ['SR_0420502_Rasshifr_Akt_P1_P1', 'SR_0420502_Rasshifr_Akt_P1_P2']
#     sheets_NKO = listSheetsName(wb, shortURLs)
#     # sheets_NKO = ['0420502 Справка о стоимости _14', '0420502 Справка о стоимости _15']
#     row_start = 11
#     for sheet in sheets_NKO:
#         ws = wb[sheet]
#         for row in range(row_start, ws.max_row):
#             correct(row, 6)
#
#     # Корректируем ИНН
#     sheet = sheetNameFromUrl(urlSheet, 'SR_0420502_Rasshifr_Akt_P7_3')
#     # sheet = '0420502 Справка о стоимости _53'
#     ws = wb[sheet]
#     row_begin = 11
#     col = 9
#     for row in range(row_begin, ws.max_row):
#         correct(row, col)


# %%
# def corrector_scha_01_14_15_53_(wb):
#     """ Убираем лишние нули (".00"), которые могут появлиться в результате анализа данных """
#
#     # формы:
#     # 0420502 Справка о стоимости чис   SR_0420502_R1
#     # 0420502 Справка о стоимости _14   SR_0420502_Rasshifr_Akt_P1_P1
#     # 0420502 Справка о стоимости _15   SR_0420502_Rasshifr_Akt_P1_P2
#     # 0420502 Справка о стоимости _53   SR_0420502_Rasshifr_Akt_P7_3
#
#     def correct(row, col):
#         # разбиваем на две части и оставляем только первую часть (без нулей)
#         ws.cell(row, col).value = str(ws.cell(row, col).value).split('.')[0]
#
#     urlSheet = codesSheets(wb)
#     # Корректируем реквизиты фонда: "№ лицензии", "ОКПО"
#     # названия вкладок
#     sheet_rekvizits = sheetNameFromUrl(urlSheet, 'SR_0420502_R1')
#     ws = wb[sheet_rekvizits]
#     correct(10, 4)  # "№ лицензии"
#     correct(10, 5)  # "ОКПО"
#
#     # Корректируем номер кредитной организации (поле:"Регистрационный номер кредитной организации")
#     # названия вкладок
#     shortURLs = ['SR_0420502_Rasshifr_Akt_P1_P1', 'SR_0420502_Rasshifr_Akt_P1_P2']
#     sheets_NKO = listSheetsName(wb, shortURLs)
#     # sheets_NKO = ['0420502 Справка о стоимости _14', '0420502 Справка о стоимости _15']
#     row_start = 11
#     for sheet in sheets_NKO:
#         ws = wb[sheet]
#         for row in range(row_start, ws.max_row):
#             correct(row, 6)
#
#     # Корректируем ИНН
#     sheet = sheetNameFromUrl(urlSheet, 'SR_0420502_Rasshifr_Akt_P7_3')
#     # sheet = '0420502 Справка о стоимости _53'
#     ws = wb[sheet]
#     row_begin = 11
#     col = 9
#     for row in range(row_begin, ws.max_row):
#         correct(row, col)


# %%
# def corrector_scha_51(wb, df_id):
#     """ Вставляем подробное описание имущества """
#     # форма:
#     # 0420502 Справка о стоимости _51   SR_0420502_Rasshifr_Akt_P7_7
#
#     """
#     Вставляем подробное описание имущества.
#     Колонка(3): 'Сведения, позволяющие определенно установить имущество'
#     """
#     # Определяем рабочий лист в конечном файле
#     wb_sheetname = sheetNameFromUrl(codesSheets(wb), 'SR_0420502_Rasshifr_Akt_P7_7')
#     # wb_sheetname = '0420502 Справка о стоимости _51'
#     wb_ws = wb[wb_sheetname]
#
#     # Определяем рабочий лист в фале-Идентификаторы
#     df_id_sheetname = 'Иное имущество'
#     df_id_ws = df_id[df_id_sheetname]
#
#     wb_row_begin = 11  # начальная строка в конечном файле
#     wb_column_id = 2  # колонка с идентификаторами в конечном файле
#     wb_column = wb_column_id + 1  # колонка с описанием имущества в конечном файле
#
#     df_row_begin = 1  # начальная строка в файле-Идентификаторы
#     df_column_id = 0  # колонка с идентификаторами в файле-Идентификаторы
#     df_column = df_column_id + 2  # колонка с описанием имущества в файле-Идентификаторы
#     df_max_row = df_id_ws.shape[0]  # количество строк в файле-Идентификаторы
#
#     # Перебираем строки с идентификаторами в конечном файле
#     for wb_row in range(wb_row_begin, wb_ws.max_row):
#         wb_cell = wb_ws.cell(wb_row, wb_column_id)
#         # перебираем ячейки в файле-Идентификатор
#         for df_row in range(df_row_begin, df_max_row):
#             df_cell = df_id_ws.loc[df_row, df_column_id]
#             # Сравниваем значения и если совпадают,
#             # то записываем в конечный файл описание имущества из файла-Идентификатор
#             if wb_cell.value == df_cell:
#                 wb_ws.cell(wb_row, wb_column).value = df_id_ws.loc[df_row_begin, df_column]
#
#     """
#     Убираем значение из итоговой строки.
#     Колонка(4): 'Иное имущество - Количество в составе активов, штук'
#     """
#     wb_ws.cell(wb_ws.max_row, 4).value = ''


# %%
def corrector_scha_51_(ws, df_identifier):
    """ Вставляем подробное описание имущества """
    # форма:
    # 0420502 Справка о стоимости _51   SR_0420502_Rasshifr_Akt_P7_7

    """ 
    Вставляем подробное описание имущества.
    Колонка(3): 'Сведения, позволяющие определенно установить имущество'
    """
    # Определяем рабочий лист в конечном файле
    # wb_sheetname = sheetNameFromUrl(codesSheets(wb), 'SR_0420502_Rasshifr_Akt_P7_7')
    # wb_sheetname = '0420502 Справка о стоимости _51'
    # wb_ws = wb[wb_sheetname]

    # Определяем рабочий лист в фале-Идентификаторы
    identifier_sheetname = 'Иное имущество'
    identifier_ws = df_identifier[identifier_sheetname]

    ws_row_begin = 11  # начальная строка в конечном файле
    ws_column_id = 2  # колонка с идентификаторами: "Вид иного имущества"
    ws_column = ws_column_id + 1  # колонка с описанием имущества в конечном файле

    identifier_row_begin = 1  # начальная строка в файле-Идентификаторы
    identifier_column_id = 0  # колонка с идентификаторами в файле-Идентификаторы
    identifier_column = identifier_column_id + 2  # колонка с описанием имущества в файле-Идентификаторы
    identifier_max_row = identifier_ws.shape[0]  # количество строк в файле-Идентификаторы

    # Перебираем строки с идентификаторами в конечном файле
    for wb_row in range(ws_row_begin, ws.max_row):
        wb_cell = ws.cell(wb_row, ws_column_id)
        # перебираем ячейки в файле-Идентификатор
        for df_row in range(identifier_row_begin, identifier_max_row):
            df_cell = identifier_ws.loc[df_row, identifier_column_id]
            # Сравниваем значения и если совпадают,
            # то записываем в конечный файл описание имущества из файла-Идентификатор
            if wb_cell.value == df_cell:
                ws.cell(wb_row, ws_column).value = identifier_ws.loc[identifier_row_begin, identifier_column]

    """ 
    Убираем значение из итоговой строки.
    Колонка(4): 'Иное имущество - Количество в составе активов, штук'
    """
    ws.cell(ws.max_row, 4).value = ''


# %%

# def corrector_scha_56to57(wb_xbrl, df_matrica, df_avancor, ERRORS, errors_file):
#     """Записываем в формы ФИО подписантов"""
#     # Формы:
#     # 0420502 Справка о стоимости _56   SR_0420502_Podpisant
#     # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep
#
#     # отбор форм подписантов
#     shortURLs = ['SR_0420502_Podpisant', 'SR_0420502_Podpisant_spec_dep']
#     sheet_2_fio = listSheetsName(wb_xbrl, shortURLs)
#     # sheet_2_fio = wb_xbrl.sheetnames[61:63]
#
#     for form in sheet_2_fio:
#         # находим, используя матрицу, раздел в файле Аванкор, соответствующий выбранной форме
#         title_1_name = df_matrica.loc[form, 'sheet_1_title']
#
#         # количество строк в таблице Аванкор
#         index_max = df_avancor.shape[0]
#         # Номер строки с названием раздела в файле Аванкор
#         title_row = razdel_name_row(df_avancor, title_1_name, index_max)
#
#         # номер колонки ячейки с данными в таблице Аванкор
#         data_col = df_matrica.loc[form, 'cell1_col']
#         data_col = column_index_from_string(data_col)
#
#         # ФИО подписанта
#         fio = df_avancor.loc[title_row, data_col]
#
#         # записываем в форму xbrl
#         ws_fio = wb_xbrl[form]
#         row_fio, col_fio = coordinate(df_matrica.loc[form, 'cell_fio'])
#         ws_fio.cell(row_fio, col_fio).value = fio


# %%
# def corrector_Podpisant_1_(ws, df_matrica, df_avancor, shortURL, ERRORS, errors_file):
#     """Записываем в формы ФИО подписантов"""
#     # Формы:
#     # 0420502 Справка о стоимости _56   SR_0420502_Podpisant
#     # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep
#
#     # отбор форм подписантов
#     # shortURLs = ['SR_0420502_Podpisant', 'SR_0420502_Podpisant_spec_dep']
#     # sheet_2_fio = listSheetsName(wb_xbrl, shortURLs)
#     # sheet_2_fio = wb_xbrl.sheetnames[61:63]
#
#     # for form in sheet_2_fio:
#
#     # находим, используя матрицу, раздел в файле Аванкор,
#     # соответствующий выбранной форме
#     title_1_name = df_matrica.loc[shortURL, 'sheet_1_title']
#
#     # количество строк в таблице Аванкор
#     index_max = df_avancor.shape[0]
#     # Номер строки с названием раздела в файле Аванкор
#     avancor_title_row = razdel_name_row(df_avancor, title_1_name, index_max)
#
#     # номер колонки ячейки с данными в таблице Аванкор
#     avancor_data_col = df_matrica.loc[shortURL, 'cell1_col']
#     avancor_data_col = column_index_from_string(avancor_data_col)
#
#     # ФИО подписанта
#     fio = df_avancor.loc[avancor_title_row, avancor_data_col]
#
#     # записываем в форму xbrl
#     # ws_fio = wb_xbrl[form]
#     row_fio, col_fio = coordinate(df_matrica.loc[shortURL, 'cell_fio'])
#     ws.cell(row_fio, col_fio).value = fio


# %%

# def corrector_Podpisant_2_(ws_xbrl, df_identifier, sheet_name, cell_with_fio):
#     """ Вставляем в ячейки с подписантами полное ФИО вместо сокращенного ФИО"""
#
#     # Формы:
#     # 0420502 Справка о стоимости _56   SR_0420502_Podpisant
#     # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep
#
#     def insert_fio_full(sheet_name, cell_with_fio, ws_xbrl):
#         """ Вставляем полное ФИО в ячейку"""
#
#         def searche_fio_full():
#             """ Ищем полное ФИО в файле с Идентификаторами """
#             fio_full = None
#             # Удаляем лишние пробелы из ФИО (если пробелы есть)
#             try:
#                 fio_short_bp = fio_short.replace(' ', '')
#
#             except AttributeError:
#                 log.error(f'форма "{sheet_name}", ФИО подписанта "{fio_short}" - не заполнено ')
#                 return fio_full
#
#             except:
#                 log.error(f'"{sheet_name}" - Неизвестная ошибка')
#                 return fio_full
#
#             # Удаляем лишние пробелы из ФИО (если пробелы есть)
#             for n in range(len(ws_identifier['ФИО_кратко'])):
#                 # n += 1
#                 ws_identifier['ФИО_кратко'][n] = ws_identifier['ФИО_кратко'][n].replace(' ', '')
#
#             try:  # находим нужное ФИО
#                 fio = ws_identifier[ws_identifier['ФИО_кратко'] == fio_short_bp]
#                 # переиндексируем df, начиная с '0'
#                 fio.index = range(0, len(fio))
#                 fio_full = fio.loc[0, 'ФИО_полностью']
#             except KeyError:
#                 log.error(f'"{sheet_name}", полное ФИО подписанта "{fio_short}" не найдено')
#
#             except:
#                 log.error(f'"{sheet_name}", Неизвестная ошибка')
#             return fio_full
#
#         # ws_xbrl = wb_xbrl[sheet_name]
#         fio_short = ws_xbrl[cell_with_fio].value
#         fio_full = searche_fio_full()
#         # Если находим полное ФИО, то вставляем его в форму xbrl
#         if fio_full:
#             ws_xbrl[cell_with_fio].value = fio_full
#             print(fio_full)
#
#     # Загружаем нужную страницу файла с Идентификаторами
#     ws_identifier = df_identifier['ФИО']
#     # Выбираем в качестве заголовкой столбцов первую строку
#     ws_identifier.columns = ws_identifier.iloc[0]
#     # удаляем нулевую строку - она дублирует заголовки
#     ws_identifier = ws_identifier.drop(ws_identifier.index[[0]])
#     # переиндексируем строки, начиная с "0"
#     ws_identifier.index -= 1
#
#     # urlSheet = codesSheets(wb_xbrl)
#     # sheet_name = sheetNameFromUrl(urlSheet, 'SR_0420502_Podpisant')
#     # sheet_name = '0420502 Справка о стоимости _56'
#     # cell_with_fio = 'B7'
#     insert_fio_full(sheet_name, cell_with_fio, ws_xbrl)


# %%
# def corrector_scha_56to57_fio_full(wb_xbrl, file_id, ERRORS, errors_file):
#     """ Вставляем в ячейки с подписантами полное ФИО вместо сокращенного ФИО"""
#
#     # Формы:
#     # 0420502 Справка о стоимости _56   SR_0420502_Podpisant
#     # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep
#
#     def insert_fio_full(sheet_name, cell_with_fio, wb_xbrl):
#         """ Вставляем полное ФИО в ячейку"""
#
#         def searche_fio_full():
#             """ Ищем полное ФИО в файле с Идентификаторами """
#             fio_full = None
#             # Удаляем лишние пробелы из ФИО (если пробелы есть)
#             try:
#                 fio_short_bp = fio_short.replace(' ', '')
#             except AttributeError:
#                 error = f'форма "{sheet_name}", ФИО подписанта "{fio_short}" - не заполнено '
#                 print(error)
#                 ERRORS.append(error)
#                 return fio_full
#
#             except:
#                 error = f'"{sheet_name}" - Неизвестная ошибка'
#                 print(error)
#                 ERRORS.append(error)
#                 return fio_full
#
#             # Удаляем лишние пробелы из ФИО (если пробелы есть)
#             for n in range(len(ws_df_id['ФИО_кратко'])):
#                 ws_df_id['ФИО_кратко'][n] = ws_df_id['ФИО_кратко'][n].replace(' ', '')
#
#             try:  # находим нужное ФИО
#                 fio = ws_df_id[ws_df_id['ФИО_кратко'] == fio_short_bp]
#                 # переиндексируем df, начиная с '0'
#                 fio.index = range(0, len(fio))
#                 fio_full = fio.loc[0, 'ФИО_полностью']
#             except KeyError:
#                 error = f'"{sheet_name}", полное ФИО подписанта "{fio_short}" не найдено'
#                 print(error)
#                 ERRORS.append(error)
#                 # write_errors(ERRORS, errors_file)
#             except:
#                 error = f'"{sheet_name}", Неизвестная ошибка'
#                 print(error)
#                 ERRORS.append(error)
#                 # write_errors(ERRORS, errors_file)
#             return fio_full
#
#         ws_xbrl = wb_xbrl[sheet_name]
#         fio_short = ws_xbrl[cell_with_fio].value
#         fio_full = searche_fio_full()
#         # Если находим полное ФИО, то вставляем его в форму xbrl
#         if fio_full:
#             ws_xbrl[cell_with_fio].value = fio_full
#             print(fio_full)
#
#     # Загружаем нужную страницу файла с Идентификаторами
#     ws_df_id = pd.read_excel(file_id, sheet_name='ФИО', header=0)
#
#     urlSheet = codesSheets(wb_xbrl)
#     sheet_name = sheetNameFromUrl(urlSheet, 'SR_0420502_Podpisant')
#     # sheet_name = '0420502 Справка о стоимости _56'
#     cell_with_fio = 'B7'
#     insert_fio_full(sheet_name, cell_with_fio, wb_xbrl)
#
#     sheet_name = sheetNameFromUrl(urlSheet, 'SR_0420502_Podpisant_spec_dep')
#     # sheet_name = '0420502 Справка о стоимости _57'
#     cell_with_fio = 'B8'
#     insert_fio_full(sheet_name, cell_with_fio, wb_xbrl)


# %%
# def corrector_scha_57(wb_xbrl, df_id, id_fond):
#     """ Проставляем id Фонда и реквизиты СпецДепа в форму"""
#     # форма:
#     # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep
#
#     # Проставляем id Фонда и реквизиты СпецДепа в форму
#     sheet_name = sheetNameFromUrl(codesSheets(wb_xbrl), 'SR_0420502_Podpisant_spec_dep')
#     ws_xbrl = wb_xbrl[sheet_name]
#     ws_id = df_id['ПИФ']
#
#     # находим строку в df 'Идентификаторы' с реквизитами СД
#     requisites = ws_id[(ws_id[0] == id_fond)]
#     # переиндексируем df, начиная с '0'
#     requisites.index = range(0, len(requisites))
#     row = 8
#
#     # вставляем идентификатор фонда
#     ws_xbrl.cell(row, 1).value = id_fond
#     # вставляем реквизиты СД
#     # номера колонок в df 'Идентификаторы' начинается с "0", а ws_xbrl с "1"
#     for col in range(3, 6):
#         ws_xbrl.cell(row, col).value = str(requisites.loc[0, col])


# %%
def corrector_Podpisant_3_(ws_xbrl, df_identifier, id_fond):
    """ Проставляем реквизиты СпецДепа"""
    # форма:
    # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep

    # Проставляем id Фонда и реквизиты СпецДепа в форму
    # sheet_name = sheetNameFromUrl(codesSheets(wb_xbrl), 'SR_0420502_Podpisant_spec_dep')
    # ws_xbrl = wb_xbrl[sheet_name]
    ws_identifier = df_identifier['ПИФ']

    # находим строку в df 'Идентификаторы' с реквизитами СД
    requisites = ws_identifier[(ws_identifier[0] == id_fond)]
    # переиндексируем df, начиная с '0'
    requisites.index = range(0, len(requisites))

    # вставляем реквизиты СД
    # номера колонок в df 'Идентификаторы' начинается с "0", а ws_xbrl с "1"
    for col in range(3, 6):
        ws_xbrl.cell(8, col).value = str(requisites.loc[0, col])


# %%

if __name__ == "__main__":
    pass
