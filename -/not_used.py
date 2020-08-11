# НЕ ИСПОЛЬЗУЕТСЯ
def find_id(ws, df_id_ws, row_i, id_i1, col_begin, id_fond):
    """ найти нужный идентификатор, анализируя содержание ячеек в строке таблицы XBRL"""
    # id_i1 - первый признак идентификатора
    # id_i2 - второй признак идентификатора

    # список найденных идентификаторов (первая колонка)
    id_list = df_id_ws[df_id_ws[1] == id_i1][0].to_list()

    id_i2 = False  # второй признак идентификатора

    # если список состоит из одного элемента, то идентификатор определен
    if len(id_list) == 1:
        id_name = id_list[0]
    else:  # идентификатор НЕ определен (в списке несколько элементов)
        print(df_id_ws[df_id_ws[1] == id_i1])
        # список вторых признаков идентификатора (третья колонка)
        id_str_2 = df_id_ws[df_id_ws[1] == id_i1][2].to_list()

        # просматриваем ячейки в строке
        for col_i in range(col_begin, ws.max_column + 1):
            # Признак идентификатора
            id_i2 = ws.cell(row_i, col_i).value

            if id_i2 in id_str_2:  # в строке есть второй признак
                # выбираем идентификатор
                id_name = df_id_ws[(df_id_ws[1] == id_i1) &
                                   (df_id_ws[2] == id_i2)][0].to_list()[0]  # .values[0]
                break

        if not id_i2:  # если в строке второй признак не найден, то находим связь с "id_fond"
            id_i2 = id_fond
            id_list = df_id_ws[(df_id_ws[1] == id_i1) & (df_id_ws[4] == id_i2)][0].to_list()
            id_name = id_list[0]

    return id_name


# НЕ ИСПОЛЬЗУЕТСЯ
def insert_id_best(ws, id_row, id_cols, df_id, rows_numbers, row_begin, col_begin, max_number, id_fond):
    """ Находим и вставляем идентификаторы в ячейки"""

    # словарь ВСЕХ названий идетификаторов (полные) и имен вкладок (сокращенные)
    list_id = {(df_id[k].loc[0, 0]): k for k in df_id.keys()}

    # список идентификаторов на листе
    id_on_list = [ws.cell(id_row, col).value for col in id_cols]
    # находим признаки идентификатора в строке
    for i, id_name in enumerate(id_on_list):  # полные названия идентификаторов
        # название листа в файле с идентификатрами
        sheet_name_id = list_id[id_name]
        # открываем нужную страницу в файле с идентификаторами
        df_id_ws = df_id[sheet_name_id]

        # находим множество всех признаков идентификаторов (исключая нулевую строку)
        # колонка [1] == (вторая колонка)
        id_n = set(df_id_ws[1].to_list()[1:])

        for row in range(len(rows_numbers) - 1):
            row_i = row_begin + row
            id_on = False
            for col in range(max_number):

                col_i = col_begin + col
                # Признак идентификатора
                id_i = ws.cell(row_i, col_i).value

                # Если признак состоит из одних цифр, то преобразуем его в целое число,
                # т.к. это либо ИНН либо ОГРН
                if type(id_i) == str and id_i.isdigit():
                    id_i = int(id_i)

                # если находим признак идентификатора, то прерываем пеербор колонок
                if id_i in id_n:
                    id_on = True
                    # записываем название идентификатора в таблицу xbrl
                    ws.cell(row_i, id_cols[i]).value = \
                        find_id(ws, df_id_ws, row_i, id_i, col_begin, id_fond)
                    break

            # исусственное установление признака "Наименование биржи" == 'NoName',
            # так как при указании на внебиржевые бумаги поле в отчетностит остается пустым
            if 'NoName' in id_n:
                id_on = True
                id_i = 'NoName'
                ws.cell(row_i, id_cols[i]).value = \
                    find_id(ws, df_id_ws, row_i, id_i, col_begin, id_fond)

            # Запоминаем ошибку если идентификатор не найден
            if not id_on:
                print('.......ERROR!.......')
                ERRORS.append(f'"{ws.title}"; Идентификатор - "{sheet_name_id}" -  не найден!')
