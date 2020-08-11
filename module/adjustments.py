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
def corrector_scha_51_(ws, df_identifier):
    """ Вставляем подробное описание имущества """
    # форма:
    # 0420502 Справка о стоимости _51   SR_0420502_Rasshifr_Akt_P7_7

    """ 
    Вставляем подробное описание имущества.
    Колонка(3): 'Сведения, позволяющие определенно установить имущество'
    """

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
def corrector_Podpisant_3_(ws_xbrl, df_identifier, id_fond):
    """ Проставляем реквизиты СпецДепа"""
    # форма:
    # 0420502 Справка о стоимости _57   SR_0420502_Podpisant_spec_dep

    # Проставляем id Фонда и реквизиты СпецДепа в форму
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
