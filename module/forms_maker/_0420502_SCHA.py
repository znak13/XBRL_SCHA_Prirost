""" Формирование форм СЧА - итоговые"""

import module.adjustments as adj
import module.functions as fun
import module.dataCopy as dcop

# from openpyxl.styles import Alignment
# from module.functions import codesSheets, listSheetsName, sheetNameFromUrl

# Формы СЧА:
# 0420502 Справка о стоимости чис	SR_0420502_R1
# 0420502 Справка о стоимости ч_2	SR_0420502_R2
# 0420502 Справка о стоимости ч_3	SR_0420502_R3_P1
# 0420502 Справка о стоимости ч_4	SR_0420502_R3_P4
# 0420502 Справка о стоимости ч_5	SR_0420502_R3_P2
# 0420502 Справка о стоимости ч_6	SR_0420502_R3_P3
# 0420502 Справка о стоимости ч_7	SR_0420502_R3_P5
# 0420502 Справка о стоимости ч_8	SR_0420502_R3_P6
# 0420502 Справка о стоимости ч_9	SR_0420502_R3_P7
# 0420502 Справка о стоимости _10	SR_0420502_R3_P8
# 0420502 Справка о стоимости _11	SR_0420502_R3_P9
# 0420502 Справка о стоимости _12	SR_0420502_R4
# 0420502 Справка о стоимости _13	SR_0420502_R5


# **********************************************************************************
def scha (wb, id_fond, df_avancor):

    # **********************************************************************************
    def copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols):
        cell_start_row, cell_start_col = fun.coordinate(AvancoreCellBegin)
        cell_end_row, cell_end_col = fun.coordinate(AvancoreCellEnd)
        # Номера колонок в таблице Аванкор (за исключением пустых)
        columns_numbers = fun.find_columns_numbers(df_avancor, cell_end_col,
                                                   AvancoreTblCols, cell_start_row,
                                                   data_col=cell_start_col)
        # Номера строк в таблице Аванкор
        rows_numbers = [x for x in range(cell_start_row, cell_end_row + 1)]
        # координаты первой ячейки в таблице xbrl
        col_begin, row_begin = fun.begin_cell(ws, AvancoreTblCols)
        # Копирование данных из таблицы Аванкор в таблицу XBRL
        dcop.copy_data(ws, df_avancor, rows_numbers, columns_numbers,
                       row_begin, col_begin)

    # **********************************************************************************
    def scha_01():
        """0420502 Справка о стоимости чис - SR_0420502_R1"""

        shortURL = 'SR_0420502_R1' # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10' # левая-верхняя ячейка с данными
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Раздел I. Реквизиты акционерного инвестиционного фонда (паевого инвестиционного фонда)'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'A10'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'M10'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 5
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Меняем местами значения ячеек и добавляем id фонда и УК
        adj.corrector_scha_01(wb, id_fond, shortURL=shortURL)
        # ---------------------------------------------------------
        # Корректируем реквизиты фонда: "№ лицензии", "ОКПО"
        ws.cell(10, 4).value = str(ws.cell(10, 4).value).split('.')[0]  # "№ лицензии"
        ws.cell(10, 5).value = str(ws.cell(10, 5).value).split('.')[0]  # "ОКПО"

        # Убираем лишние '.00' в конце строки,
        # которые могут появиться после копирования:
        # "номер лицензии", "ОКПО"
        adj.corrector_00_v2(ws, 'D', 'E', row=10)

    # **********************************************************************************
    def scha_02():
        """ 0420502 Справка о стоимости ч_2	SR_0420502_R2"""

        shortURL = 'SR_0420502_R2'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'B9'    # левая-верхняя ячейка с данными
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Раздел II. Параметры справки о стоимости чистых активов'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'A18'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'G18'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 3
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)


    # **********************************************************************************
    def scha_03():
        """0420502 Справка о стоимости ч_3	SR_0420502_R3_P1"""
        shortURL = 'SR_0420502_R3_P1'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'   # левая-верхняя ячейка с данными
        cellPeriod = 'C7'   # ячейка с периодом отчетности
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 1. Денежные средства'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I25'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O31'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_04():
        """0420502 Справка о стоимости ч_4	SR_0420502_R3_P4"""
        shortURL = 'SR_0420502_R3_P4'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 4. Недвижимое имущество и права аренды недвижимого имущества'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I72'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O78'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_05():
        """0420502 Справка о стоимости ч_5	SR_0420502_R3_P2"""
        shortURL = 'SR_0420502_R3_P2'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 2. Ценные бумаги российских эмитентов (за исключением закладных)'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I37'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O52'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_06():
        """0420502 Справка о стоимости ч_6	SR_0420502_R3_P3"""
        shortURL = 'SR_0420502_R3_P3'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 3. Ценные бумаги иностранных эмитентов'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I58'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O66'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_07():
        """0420502 Справка о стоимости ч_7	SR_0420502_R3_P5"""
        shortURL = 'SR_0420502_R3_P5'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 5. Имущественные права (за исключением прав аренды недвижимого имущества, прав из кредитных договоров и договоров займа и прав требования к кредитной организации выплатить денежный эквивалент драгоценных металлов)'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I84'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O89'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_08():
        """0420502 Справка о стоимости ч_8	SR_0420502_R3_P6"""
        shortURL = 'SR_0420502_R3_P6'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 6. Денежные требования по кредитным договорам и договорам займа, в том числе удостоверенные закладными'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I95'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O97'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_09():
        """0420502 Справка о стоимости ч_9	SR_0420502_R3_P7"""
        shortURL = 'SR_0420502_R3_P7'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 7. Иное имущество, не указанное в подразделах 1-6'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I103'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O111'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

        # ---------------------------------------------------------
        # ДОПОЛНИТЕЛЬНО:
        # в случае формирования формы:
        # 0420502 Справка о стоимости _51	SR_0420502_Rasshifr_Akt_P7_7
        # '7.7. Иное имущество, не указанное в таблицах пунктов 7.1 - 7.6'
        # проставляем нули в первой и последних строках таблицы
        # (если в этих ячейках не было данных)

    # **********************************************************************************
    def scha_10():
        """0420502 Справка о стоимости _10	SR_0420502_R3_P8"""
        shortURL = 'SR_0420502_R3_P8'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 8. Дебиторская задолженность'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I117'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O121'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Проставляем нулевые значения в форме
        adj.corrector_scha_03to10(wb, shortURL=shortURL)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

    # **********************************************************************************
    def scha_11():
        """0420502 Справка о стоимости _11	SR_0420502_R3_P9"""
        shortURL = 'SR_0420502_R3_P9'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'B10'
        cellPeriod = 'B6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Подраздел 9. Общая стоимость активов'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I127'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'M127'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 3
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)
        # sheetName = sheetNameFromUrl(codesSheets(wb), shortURL)
        # ws = wb[sheetName]
        # for col in range(2, ws.max_column + 1):
        #     ws.cell(10, col).alignment = Alignment(horizontal='right')
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)

    # **********************************************************************************
    def scha_12():
        """0420502 Справка о стоимости _12	SR_0420502_R4"""
        shortURL = 'SR_0420502_R4'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C10'
        cellPeriod = 'C6'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Раздел IV. Обязательства'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I133'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'O137'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 4
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Вставляем данные о периоде отчетности
        adj.corrector_scha_03to12_(wb, df_avancor, cellPeriod, shortURL=shortURL)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)
        # for row in range(10, ws.max_row + 1):
        #     for col in range(3, ws.max_column + 1):
        #         ws.cell(row, col).alignment = Alignment(horizontal='right')
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)

    # **********************************************************************************
    def scha_13():
        """0420502 Справка о стоимости _13	SR_0420502_R5"""
        shortURL = 'SR_0420502_R5'  # код вкладки
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        cellBegin = 'C9'
        print(f'{sheetName} - {shortURL}')
        # ---------------------------------------------------------
        # Переносим данные в форму:
        # Заголовки формы в файле-Аванкор
        AvancoreTitle = 'Раздел V. Стоимость чистых активов'
        # Первая ячейка с данными (левая-верхняя)
        AvancoreCellBegin = 'I143'
        # Последняя ячейка (правая-нижняя)
        AvancoreCellEnd = 'K145'
        # Количество колонок для копирования в таблице Аванкор
        AvancoreTblCols = 2
        # Копируем данные из файла Аванкор в форму XBRL
        copyFromAvancore(ws, AvancoreCellBegin, AvancoreCellEnd, AvancoreTblCols)
        # ---------------------------------------------------------
        # Корректируем количество паев,
        # устанавливая нужную точностью знаков после запятой
        adj.corrector_scha_13_(id_fond, ws, df_avancor)
        # ---------------------------------------------------------
        # Записываем в форму идентификатор фонда
        dcop.copy_id_fond_to_tbl(ws, id_fond)
        # ---------------------------------------------------------
        # Форматируем ячейки
        fun.cellFormat(ws, cellBegin)

        # **********************************************************************************

    urlSheets = fun.codesSheets(wb)  # словарь - "код вкладки":"имя вкладки"
    scha_01()
    scha_02()
    scha_03()
    scha_04()
    scha_05()
    scha_06()
    scha_07()
    scha_08()
    scha_09()
    scha_10()
    scha_11()
    scha_12()
    scha_13()



# %%

if __name__ == "__main__":
    pass