from module.dataCheck   import check_errors
from module.analiz_data import analiz_data_all
from module.functions   import coordinate, find_columns_numbers, \
    razdel_name_row, start_data_row, end_data_row, listSheetsName


def copy_data(ws, df_avancor, rows_numbers, columns_numbers, row_begin, col_begin, ERRORS):
    """ Копирование данных из таблицы Аванкор в таблицу XBRL """

    for r, row in enumerate(rows_numbers):
        for c, col in enumerate(columns_numbers):
            cell_avancor = df_avancor.loc[row, col]
            check_errors(ws, cell_avancor, row_begin, r, c, col_begin, ERRORS )
            if cell_avancor != '-' and \
                    cell_avancor != "x" and \
                    cell_avancor != 0 and \
                    str(cell_avancor) != 'nan':
                ws.cell(row_begin + r, col_begin + c).value = \
                    analiz_data_all(df_avancor.loc[row, col])

#%%
def copy_id_fond_to_tbl(ws, id_fond):
    """Записываем в форму идентификатор фонда"""

    # идентификатор фонда содержится во всех формах, кроме:
    # 0420502 Справка о стоимости чис - SR_0420502_R1
    # 0420502 Справка о стоимости _56 - SR_0420502_Podpisant
    # 0420502 Справка о стоимости _57 - SR_0420502_Podpisant_spec_dep

    # Во всех формах идентификатор фонда содержится
    # в строке "5" и в крайней правой колонке,
    # а текст в ячейке начинается с 'Z= Идентификатор АИФ ПИФ-'

    cell_id = ws.cell(row=5, column=ws.max_column)
    info = 'Z= Идентификатор АИФ ПИФ-'
    if info in cell_id.value:
        cell_id.value = info + id_fond
#%%
def id_serch(ws, sheet_name_id, row_i, col_begin, id_fond, df_id, ERRORS):
    """ Определение идентификатора
        :ws: форма в отчетности xbrl
        :sheet_name_id: название листа с идентификатором
        :row: текущая строка в в таблице xbrl
        :col_begin: начальная колонка в таблице xbrl
        :id_fond: идентификатор фонда
        :return 'id_is': искомый идентификатор
    """

    def round_id(df_id_ws, n, id_fond=None):

        for col in range(col_begin, ws.max_column + 1):
            id_priznak = ws.cell(row_i, col).value if id_fond is None else id_fond

            # Список признаков
            df_priznak = df_id_ws[df_id_ws[n] == id_priznak]

            # Если находим признак в строке, то прерываем перебор колонок
            if id_priznak in df_priznak[n].to_list():
                break

        return df_priznak, id_priznak

    df_id_ws = df_id[sheet_name_id]
    # Меняем тип всех данных на "str"
    # (это необходимо для корректного сравнения в дальнейшем)
    df_id_ws = df_id_ws.astype(str)

    df_priznak = df_id_ws  # текущий df
    df_priznak_list = []  # список найденных df
    id_is = None  # искомый идентификатор

    n = 1
    while n < df_priznak.columns.size:
        df_priznak, id_priznak = round_id(df_priznak, n)
        # допоминаем список c 'df_priznak'
        df_priznak_list.append(df_priznak)
        if df_priznak.index.size == 1:  # одна строчка
            # идентификатор - это превый элемент в 'df_priznak'
            id_is = df_priznak[0].values[0]
            break
        elif df_priznak.index.size == 0:  # если список признаков пуст, то ищем по 'id_fond'
            if sheet_name_id == 'Биржа':  # идентификатор не найден
                id_is = df_id_ws.loc[1, 0]
            elif sheet_name_id == 'Банковский счет':
                # берем предыдущий 'df_priznak' () из 'df_priznak_list'
                # (отнимаем "2",т.к. в списке 'df_priznak_list' элементф начинаются с "0")
                df_priznak, id_priznak = round_id(df_priznak_list[n - 2], n + 1, id_fond=id_fond)
                # идентификатор - это превый элемент в 'df_priznak'
                try:
                    id_is = df_priznak[0].values[0]
                except Exception:
                    ERRORS.append(f'ERROR! - в файле "Идентификаторы" не указан расчетный счет фонда')
            else:
                id_is = 'ошибка'
                print('------>ERROR!', round_id.__name__)
                ERRORS.append(f'ERROR! - в форме "{ws.title}" не определен "{df_id_ws.loc[0, 0]}", строка:{row_i} ')

            break
        else:
            n += 1

    if id_is:
        return id_is
    else:
        id_is = 'ошибка'
        print('.......ERROR!.......')
        print('Идентификатор не найден')
        ERRORS.append(f'Идентификатор не найден: '
                      f'"{ws.title}", "{df_id_ws.loc[0, 0]}", строка:{row_i}')

        return id_is
#%%

def insert_id(ws, id_row, id_cols, df_id, rows_numbers, row_begin, col_begin, id_fond, ERRORS):
    """ Находим и вставляем идентификаторы в ячейки"""

    # словарь ВСЕХ названий идетификаторов (полные) и имен вкладок (сокращенные)
    list_id = {(df_id[k].loc[0, 0]): k for k in df_id.keys()}
    # если ошибка в этой строке, то проверить названеи вкладок в файле с идентификаторами
    # ...

    # список идентификаторов на листе
    id_on_list = [ws.cell(id_row, col).value for col in id_cols]
    # находим признаки идентификатора в строке
    for i, id_name in enumerate(id_on_list):  # полные названия идентификаторов
        # название листа в файле с идентификатрами (сокращенные названия идентификаторов)
        sheet_name_id = list_id[id_name]

        for row in range(len(rows_numbers) - 1):
            row_i = row_begin + row
            id_is = id_serch(ws, sheet_name_id, row_i, col_begin, id_fond, df_id, ERRORS)

            # записываем название идентификатора в таблицу xbrl
            ws.cell(row_i, id_cols[i]).value = id_is
#%%

def copy_zapiski(wb, df_matrica, df_avancor, urlSheets, ERRORS):

    # кол-во строк и столбцов в файле Аванкор
    index_max = df_avancor.shape[0]
    collumn_max = df_avancor.shape[1]

    zapiski_null = []  # список пустых форм
    for url,form in urlSheets.items():

        print(f'{form}')
        ws = wb[form]
        # находим, используя матрицу, раздел в файле-Аванкор, соответствующий выбранной форме
        title_1_name = df_matrica.loc[url, 'sheet_1_title']

        if title_1_name == title_1_name:
            # если ячейка в столбце "sheet_1_title" пустая, то "title_1_name = nan"
            # в этом случае: bool(title_1_name == title_1_name) == Fals
            # (странно, но работает)

            # Номер строки с названием раздела в файле Аванкор"""
            title_row = razdel_name_row(df_avancor, title_1_name, index_max)

            # находим номер первой строку с данными в файле Аванкор
            data_row = start_data_row(df_avancor, index_max, title_row)

            # находим номер последней строки с данными в таблице Аванкор"""
            row_end = end_data_row(df_avancor, index_max, data_row)

            if df_avancor.loc[row_end - 1, 3] != '2':
                # список всех номеров строк в таблице Аванкор
                rows_numbers = [x for x in range(data_row, row_end)]
            else:
                # Если в соседней ячейки(колонка "C") == "2", то это заголовок таблицы
                # и, следовательно, форма пустая
                # Запоминаем название вкладки
                zapiski_null.append(form)
                # переходим к поиску следующего раздела
                continue

            # цифра в последней колонке и номер строки с идектификаторами в таблице XBRL
            # (количество колонок для копирования)
            max_number = df_matrica.loc[url, 'tbl_col']

            # координаты первой ячейки в таблице XBRL
            cell_start = df_matrica.loc[url, 'cell2']
            cell_start_row, cell_start_col = coordinate(cell_start)

            # Номера колонок в таблице Аванкор, за исключением пустых
            columns_numbers = find_columns_numbers(df_avancor, collumn_max, max_number, data_row, data_col=3)

            # Копирование данных из таблицы Аванков в таблицу XBRL
            copy_data(ws, df_avancor, rows_numbers, columns_numbers, cell_start_row, cell_start_col, ERRORS)
            ERRORS.append(f'"{ws.title}" - вставьте "Идентификатор строки"')

        else:
            # "title_1_name = nan" - в файле-Аванкор нет раздела, соответствующего этой форме
            zapiski_null.append(form)

    return zapiski_null