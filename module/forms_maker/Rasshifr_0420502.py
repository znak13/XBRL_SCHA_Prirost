""" Формирование форм СЧА - расшифровки разделов"""

from module.functions import listSheetsName
import module.functions as fun
import module.dataCopy as dcop
import module.adjustments as adj

# Выбираем формы: расшифровки разделов
# 0420502 Справка о стоимости _14	SR_0420502_Rasshifr_Akt_P1_P1
# 0420502 Справка о стоимости _15	SR_0420502_Rasshifr_Akt_P1_P2
# 0420502 Справка о стоимости _16	SR_0420502_Rasshifr_Akt_P2_1
# 0420502 Справка о стоимости _17	SR_0420502_Rasshifr_Akt_P2_10
# 0420502 Справка о стоимости _18	SR_0420502_Rasshifr_Akt_P2_11
# 0420502 Справка о стоимости _19	SR_0420502_Rasshifr_Akt_P2_2
# 0420502 Справка о стоимости _20	SR_0420502_Rasshifr_Akt_P2_3
# 0420502 Справка о стоимости _21	SR_0420502_Rasshifr_Akt_P2_4
# 0420502 Справка о стоимости _22	SR_0420502_Rasshifr_Akt_P2_5
# 0420502 Справка о стоимости _23	SR_0420502_Rasshifr_Akt_P2_6
# 0420502 Справка о стоимости _24	SR_0420502_Rasshifr_Akt_P2_7
# 0420502 Справка о стоимости _25	SR_0420502_Rasshifr_Akt_P2_8
# 0420502 Справка о стоимости _26	SR_0420502_Rasshifr_Akt_P2_9
# 0420502 Справка о стоимости _27	SR_0420502_Rasshifr_Akt_P3_1
# 0420502 Справка о стоимости _28	SR_0420502_Rasshifr_Akt_P3_3
# 0420502 Справка о стоимости _29	SR_0420502_Rasshifr_Akt_P3_4
# 0420502 Справка о стоимости _30	SR_0420502_Rasshifr_Akt_P3_5
# 0420502 Справка о стоимости _31	SR_0420502_Rasshifr_Akt_P3_6
# 0420502 Справка о стоимости _32	SR_0420502_Rasshifr_Akt_P3_7
# 0420502 Справка о стоимости _33	SR_0420502_Rasshifr_Akt_P4_1
# 0420502 Справка о стоимости _34	SR_0420502_Rasshifr_Akt_P4_2_1
# 0420502 Справка о стоимости _35	SR_0420502_Rasshifr_Akt_P4_2_2
# 0420502 Справка о стоимости _36	SR_0420502_Rasshifr_Akt_P5_1
# 0420502 Справка о стоимости _37	SR_0420502_Rasshifr_Akt_P5_2
# 0420502 Справка о стоимости _38	SR_0420502_Rasshifr_Akt_P5_3
# 0420502 Справка о стоимости _39	SR_0420502_Rasshifr_Akt_P5_4
# 0420502 Справка о стоимости _40	SR_0420502_Rasshifr_Akt_P5_5
# 0420502 Справка о стоимости _41	SR_0420502_Rasshifr_Akt_P6_1_1
# 0420502 Справка о стоимости _42	SR_0420502_Rasshifr_Akt_P6_1_2
# 0420502 Справка о стоимости _43	SR_0420502_Rasshifr_Akt_P6_2_1
# 0420502 Справка о стоимости _44	SR_0420502_Rasshifr_Akt_P6_2_2
# 0420502 Справка о стоимости _45	SR_0420502_Rasshifr_Akt_P7_1
# 0420502 Справка о стоимости _46	SR_0420502_Rasshifr_Akt_P7_2
# 0420502 Справка о стоимости _47	SR_0420502_Rasshifr_Akt_P7_3
# 0420502 Справка о стоимости _48	SR_0420502_Rasshifr_Akt_P7_4
# 0420502 Справка о стоимости _49	SR_0420502_Rasshifr_Akt_P7_5
# 0420502 Справка о стоимости _50	SR_0420502_Rasshifr_Akt_P7_6
# 0420502 Справка о стоимости _51	SR_0420502_Rasshifr_Akt_P7_7
# 0420502 Справка о стоимости _52	SR_0420502_Rasshifr_Akt_P8_1
# 0420502 Справка о стоимости _53	SR_0420502_Rasshifr_Akt_P8_2
# 0420502 Справка о стоимости _54	SR_0420502_Rasshifr_Ob_P1
# 0420502 Справка о стоимости _55	SR_0420502_Rasshifr_Ob_P2
# 0420502 Справка о стоимости _58	SR_0420502_Rasshifr_Akt_P3_2

def urlForms():
    return  ['SR_0420502_Rasshifr_Akt_P1_P1',   #0
            'SR_0420502_Rasshifr_Akt_P1_P2',   #1
            'SR_0420502_Rasshifr_Akt_P2_1',   #2
            'SR_0420502_Rasshifr_Akt_P2_10',   #3
            'SR_0420502_Rasshifr_Akt_P2_11',    #4
            'SR_0420502_Rasshifr_Akt_P2_2',    #5
            'SR_0420502_Rasshifr_Akt_P2_3',   #6
            'SR_0420502_Rasshifr_Akt_P2_4',   #7
            'SR_0420502_Rasshifr_Akt_P2_5',   #8
            'SR_0420502_Rasshifr_Akt_P2_6',   #9
            'SR_0420502_Rasshifr_Akt_P2_7',   #10
            'SR_0420502_Rasshifr_Akt_P2_8',   #11
            'SR_0420502_Rasshifr_Akt_P2_9',   #12
            'SR_0420502_Rasshifr_Akt_P3_1',   #13
            'SR_0420502_Rasshifr_Akt_P3_3',   #14
            'SR_0420502_Rasshifr_Akt_P3_4',   #15
            'SR_0420502_Rasshifr_Akt_P3_5',   #16
            'SR_0420502_Rasshifr_Akt_P3_6',   #17
            'SR_0420502_Rasshifr_Akt_P3_7',   #18
            'SR_0420502_Rasshifr_Akt_P4_1',   #19
            'SR_0420502_Rasshifr_Akt_P4_2_1',   #20
            'SR_0420502_Rasshifr_Akt_P4_2_2',   #21
            'SR_0420502_Rasshifr_Akt_P5_1',   #22
            'SR_0420502_Rasshifr_Akt_P5_2',   #23
            'SR_0420502_Rasshifr_Akt_P5_3',   #24
            'SR_0420502_Rasshifr_Akt_P5_4',   #25
            'SR_0420502_Rasshifr_Akt_P5_5',   #26
            'SR_0420502_Rasshifr_Akt_P6_1_1',   #27
            'SR_0420502_Rasshifr_Akt_P6_1_2',   #28
            'SR_0420502_Rasshifr_Akt_P6_2_1',   #29
            'SR_0420502_Rasshifr_Akt_P6_2_2',   #30
            'SR_0420502_Rasshifr_Akt_P7_1',   #31
            'SR_0420502_Rasshifr_Akt_P7_2',   #32
            'SR_0420502_Rasshifr_Akt_P7_3',   #33
            'SR_0420502_Rasshifr_Akt_P7_4',   #34
            'SR_0420502_Rasshifr_Akt_P7_5',   #35
            'SR_0420502_Rasshifr_Akt_P7_6',   #36
            'SR_0420502_Rasshifr_Akt_P7_7',   #37
            'SR_0420502_Rasshifr_Akt_P8_1',   #38
            'SR_0420502_Rasshifr_Akt_P8_2',   #39
            'SR_0420502_Rasshifr_Ob_P1',   #40
            'SR_0420502_Rasshifr_Ob_P2',   #41
            'SR_0420502_Rasshifr_Akt_P3_2']   #42

#%%
def scha_rashifr_All(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS):
    """ Копируем данные форм-расшифровок"""

    # Коды вкладок
    shortUrls = urlForms()
    # кол-во строк и столбцов в файле Аванкор
    index_max = df_avancor.shape[0]
    collumn_max = df_avancor.shape[1]

    codes = fun.codesSheets(wb)

    form_null = []  # список пустых форм

    for url in shortUrls:
        formName = fun.sheetNameFromUrl(codes, url)
        print(f'{formName}')
        ws = wb[formName]

        # находим, используя матрицу, раздел в файле Аванкор, соответствующий выбранной форме
        title_1_name = df_matrica.loc[url, 'sheet_1_title']

        # Номер строки с названием раздела в файле Аванкор"""
        title_row = fun.razdel_name_row(df_avancor, title_1_name, index_max)

        # находим номер первой строку с данными в файле Аванкор
        data_row = fun.start_data_row(df_avancor, index_max, title_row)

        # находим номер последнке строки с данными в таблице Аванкор"""
        row_end = fun.end_data_row(df_avancor, index_max, data_row)

        if data_row != row_end:
            # список всех номеров строк в таблице Аванкор
            rows_numbers = [x for x in range(data_row, row_end + 1)]
        else:
            # если номера первой и последней строки совпадают, то Раздел пуст и
            # запоминаем название вкладки
            form_null.append(formName)
            # переходим к поиску следующего раздела
            continue

        # цифра в последней колонке и номер строки с идектификаторами в таблице XBRL
        # (количество колонок для копирования)
        max_number, id_row = fun.end_col_number(ws)

        # Вставить нужное кол-во строк перед строкой "Итого"
        ws.insert_rows(ws.max_row, amount=(len(rows_numbers) - 1))

        # координаты первой ячейки в таблице XBRL
        col_begin, row_begin = fun.begin_cell(ws, max_number)

        # номера колонок с иденитификаторами в таблице XBRL
        id_cols = fun.id_nombers(col_begin)

        # Номера колонок в таблице Аванкор, за исключением пустых"""
        columns_numbers = fun.find_columns_numbers(df_avancor, collumn_max, max_number, data_row)

        # Копирование данных из таблицы Аванков в таблицу XBRL
        dcop.copy_data(ws, df_avancor, rows_numbers, columns_numbers, row_begin, col_begin, ERRORS)

        # Копирование идентификаторов
        # Находим и вставляем идентификаторы в ячейки
        dcop.insert_id(ws, id_row, id_cols, df_identifier, rows_numbers, row_begin, col_begin, id_fond, ERRORS)

    return form_null

def scha_rashifr(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS):

    urlSheets = fun.codesSheets(wb)

    # Переносим данные во все формы
    # и возвращаем список пустых форм
    sheetsNameNull = scha_rashifr_All(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS)

    # ================================================================
    # Далее корректируем формы в зависимости от их особенностей:
    # ================================================================
    """0420502 Справка о стоимости _14 - SR_0420502_Rasshifr_Akt_P1_P1"""
    shortURL = urlForms()[0]
    sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)
    ws = wb[sheetName]
    # Корректируем номер кредитной организации
    # (поле:"Регистрационный номер кредитной организации")
    # начиная с 11 строки
    for row in range(11, ws.max_row):
        ws.cell(row, 6).value = str(ws.cell(row, 6).value).split('.')[0]

    # ================================================================
    """0420502 Справка о стоимости _15 - SR_0420502_Rasshifr_Akt_P1_P2"""
    shortURL = urlForms()[1]
    sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)
    ws = wb[sheetName]
    # Корректируем номер кредитной организации
    # (поле:"Регистрационный номер кредитной организации")
    # начиная с 11 строки, колонка 6
    for row in range(11, ws.max_row):
        ws.cell(row, 6).value = str(ws.cell(row, 6).value).split('.')[0]

    # ================================================================
    """0420502 Справка о стоимости _51	SR_0420502_Rasshifr_Akt_P7_7"""
    shortURL = urlForms()[37]
    sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)
    ws = wb[sheetName]

    # Добавляем подробное описание имущества
    adj.corrector_scha_51_(ws, df_identifier)

    # ================================================================
    """0420502 Справка о стоимости _53 - SR_0420502_Rasshifr_Akt_P8_2"""
    shortURL = urlForms()[39]
    sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)
    ws = wb[sheetName]
    # Корректируем ИНН
    # начиная с 11 строки, колонка 9
    for row in range(11, ws.max_row):
        ws.cell(row, 9).value = str(ws.cell(row, 9).value).split('.')[0]

    # ================================================================
    """ Удаляем пустые формы-расшифровки"""
    for name in sheetsNameNull:
        wb.remove(wb[name])

# %%

if __name__ == "__main__":

    # Загружвем данные
    import module.dataLoad as ld
    id_fond, file_id, df_identifier, df_avancor, df_matrica, wb, \
    file_fond_name, errors_file = ld.load_data_2()
    ERRORS = []
    #----------------------------------------------------------------

    scha_rashifr(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS)

    # Сохраняем результат
    wb.save(file_fond_name)
