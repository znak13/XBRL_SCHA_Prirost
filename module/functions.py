import sys
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.utils import column_index_from_string


# %%
def write_errors(ERRORS, errors_file):
    """ Записываем ошибки в файл"""

    if not ERRORS:
        ERRORS.append('Ошибок не выявлено!')
    with open(errors_file, "w") as file:
        for k in ERRORS:
            file.write(str(k) + '\n\n')


# %%
def coordinate(cell):
    """Конвртер координат: 'A10' преобразуем в '10, 1' """
    data = coordinate_from_string(cell)
    row = data[1]
    col = column_index_from_string(data[0])
    return row, col


# %%
def end_col_number(ws, row_id_fond=5):
    """ Поиск номера столбца в последней колонке таблицы XBRL """
    # row_id_fond - номер строки, в которой расположенг идентификатор фонда

    max_number = row = 0
    # номер последней колонки
    for row in range(1, 5):
        cell = ws.cell(row_id_fond + row, ws.max_column)
        if cell.value and cell.value.isdigit():  # если в ячейке цифра, то заканчиваем цикл
            max_number = int(cell.value)
            break
    if not max_number or not row:
        print(f'Ошибка в функции: "end_col_number()"')

    return max_number, row_id_fond + row + 1  # строка с идентификаторами ниже на одну строчку


# %%
def begin_cell(ws, max_number, row_id_fond=5):
    """ Поиск координат первой ячейки в таблице XBRL """
    # max_number - цифра в ячейке правой крайней колонки

    row_begin = 0
    col_begin = ws.max_column - (max_number - 1)
    for row in range(3, 10):
        cell = ws.cell(row_id_fond + row, col_begin)
        if cell.fill.start_color.index == '00000000':
            row_begin = row_id_fond + row
            break
    if not row_begin:
        print(f'.......... Ошибка в функции: "begin_cell()"')

    return col_begin, row_begin


# %%
def id_nombers(col_begin):
    """ Поиск номеров колонок с иденитификаторами в таблице XBRL"""

    # номер колонкс с "Итого"
    col_itogo = 2
    # количество колонок с идентификаторами
    id_numbers = col_begin - col_itogo

    # номера колонок с идентификаторами
    id_cols = [(col_itogo + col) for col in range(id_numbers)]

    return id_cols


# %%
def razdel_name_row(df_avancor, title_1_name, index_max, ERRORS, errors_file, title_col=2):
    """Поиск Номера строки с названием раздела в файле Аванкор"""

    for row in range(1, index_max):
        title = str(df_avancor.loc[row, title_col])
        if title == title_1_name or title[:20] == title_1_name[:20]:
            title_row = row
            return title_row

    print('------>ERROR!', razdel_name_row.__name__)

    ERRORS.append(f'Раздел отчетности: "{title_1_name}" в таблице-Аванкор не найден')
    write_errors(ERRORS, errors_file)
    sys.exit("Ошибка!")


# %%
def start_data_row(df_avancor, index_max, title_row, title_col=2):
    """Поиск номера первой строки с данными в отдельной таблице Аванкор"""
    # index_max - кол-во строк в файле
    # title_col - номер колонки с названием таблицы
    # title_row - номер строки с названием таблицы

    for row in range(1, index_max):
        data_i = df_avancor.loc[title_row + row, title_col]
        if str(data_i) == '1':
            data_row = title_row + row + 1
            return data_row


# %%
def end_data_row(df_avancor, index_max, data_row, title_col=2):
    """Поиск номера последней строки с данными в отдельной таблице Аванкор"""
    # index_max - кол-во строк в файле
    # data_row  - номер первой строки с данными
    # title_col - номер колонки с названием таблицы

    data = df_avancor.loc[data_row, title_col]
    if data == 'Итого' or str(data) == 'nan' or str(data).startswith('Оценочная стоимость'):
        # print(' - раздел пуст')
        row_end = data_row
    else:
        for row in range(1, index_max):
            data_i = df_avancor.loc[data_row + row, title_col]
            if data_i == 'Итого' or str(data_i) == 'nan':
                row_end = data_row + row
                break

    return row_end


# %%
def find_columns_numbers(df_avancor, collumn_max, max_number, data_row, data_col=3):
    """ Поиск номеров колонок в таблице Аванкор, за исключением пустых"""
    # data_row, data_col_1 - координаты первой колонки с данными в Аванкор
    # max_number - номер последней колонки в таблице XBRL

    column = []
    for i in range(0, max_number):
        while str(df_avancor.loc[data_row - 1, data_col]) == 'nan' and data_col <= collumn_max:
            data_col += 1
        column.append(data_col)
        data_col += 1
    return column

# %%

def codesSheets(wb):
    """ Список кодов и наименования листов """
    codes_sheets = {}
    for sheet in wb.sheetnames:
        ws_xbrl = wb[sheet]
        ws_xbrl_cell = ws_xbrl.cell(3, 1)
        codes_sheets[ws_xbrl_cell.value] = sheet
        # print(sheet, ws_xbrl_cell.value)

    # исключаем из списка вкладку '_dropDownSheet'
    for code in codes_sheets:
        if codes_sheets[code] == '_dropDownSheet':
            codes_sheets.pop(code)
            break

    return codes_sheets

def sheetsCodes(wb):
    """ Список наименования листов и их URL  """
    sheets_codes = {}
    for sheet in wb.sheetnames:
        ws_xbrl = wb[sheet]
        ws_xbrl_cell = ws_xbrl.cell(3, 1)
        sheets_codes [sheet] = ws_xbrl_cell.value
        # print(sheet, ws_xbrl_cell.value)

    # исключаем из списка вкладку '_dropDownSheet'
    for sheet in sheets_codes:
        if sheet == '_dropDownSheet':
            sheets_codes.pop(sheet)
            break

    return sheets_codes
# %%


if __name__ == "__main__":
    import openpyxl

    file_name = r"d:\Clouds\YandexDisk\Git\XBRL_SCHA_Prirost\0420502_0420503.xlsx"
    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_name)

    codes_sheets = codesSheets(wb)
    sheets_codes = sheetsCodes(wb)

    code_PZ_5 = 'SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch_2'

    sheets_PZ=[]
    for code in codes_sheets:
        if ("_PZ_" in code) and (code_PZ_5 not in code):
            sheets_PZ.append(codes_sheets[code])

    sheets_PZ2=[]
    for sheet in sheets_codes:
        code = sheets_codes[sheet]
        if ("_PZ_" in code) and (code_PZ_5 not in code):
            sheets_PZ2.append(sheet)

