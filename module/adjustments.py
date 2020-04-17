""" Корректировки содержания форм отчетности после копирования данных из файла Аванкор"""

from openpyxl.utils import column_index_from_string
import pandas as pd

from module.functions import razdel_name_row, coordinate, write_errors
from module.analiz_data import analiz_data_data, toFixed


def corrector_scha_ch_1(wb, df_id, id_fond):
    """ Меняем местами значения ячеек """
    # форма:
    # 0420502 Справка о стоимости чис

    """ Меняем значения ячеек, т.к. в таблице xbrl ячейки расположены в ином порядке,
    чем в файле, созданном Аванкор """
    ws_change = wb['0420502 Справка о стоимости чис']

    ws_change.cell(10, 5).value, ws_change.cell(10, 6).value = \
        ws_change.cell(10, 6).value, ws_change.cell(10, 5).value

    ws_change.cell(10, 5).value, ws_change.cell(10, 7).value = \
        ws_change.cell(10, 7).value, ws_change.cell(10, 5).value

    """ Записываем в форму идентификаторы: фонда и УК """
    ws_change.cell(10, 1).value = id_fond
    ws_change.cell(10, 2).value = df_id['УК АИФ ПИФ'].loc[1, 0]


# %%
def corrector_scha_ch_3to10(wb):
    """ Проставляем нулевые значения в формах: ч_3 - ч_10 """

    # Формы:
    # '0420502 Справка о стоимости ч_3',
    # '0420502 Справка о стоимости ч_4',
    # '0420502 Справка о стоимости ч_5',
    # '0420502 Справка о стоимости ч_6',
    # '0420502 Справка о стоимости ч_7',
    # '0420502 Справка о стоимости ч_8',
    # '0420502 Справка о стоимости ч_9',
    # '0420502 Справка о стоимости _10'

    def insert_00(ws):
        """ Проставляем нулевые значения в форме """

        def perebor(cell: str, col: int):
            """ Перебор колонок и проставление нулей """
            if cell:  # есть данные
                if not ws.cell(row, col + 2).value:  # в колонке с 'долей' нет данных
                    ws.cell(row, col + 2).value = '0.00'
            else:
                ws.cell(row, col).value = '0.00'
                ws.cell(row, col + 2).value = '0.00'

        row_begin = 10
        col_1 = 3  # номер первой колонки с данными

        for row in range(row_begin, ws.max_row + 1):
            cell_1 = ws.cell(row, col_1).value
            cell_2 = ws.cell(row, col_1 + 1).value
            cell_3 = ws.cell(row, col_1 + 2).value
            cell_4 = ws.cell(row, col_1 + 3).value

            if cell_1 or cell_2:  # Если есть данные в одной из ячеек
                if not cell_1:
                    ws.cell(row, col_1).value = '0.00'
                if not cell_2:
                    ws.cell(row, col_1 + 1).value = '0.00'
                if not cell_3:
                    ws.cell(row, col_1 + 2).value = '0.00'
                if not cell_4:
                    ws.cell(row, col_1 + 3).value = '0.00'

    # Выбираем нужные формы
    sheets = wb.sheetnames[8:16]
    # проставляем нули
    for sheet in sheets:
        ws = wb[sheet]
        insert_00(ws)


# %%
def corrector_scha_ch_2to12(wb_xbrl, df_avancor, df_matrica):
    """ Вставляем данные о периоде отчетности """
    # Формы:
    # 0420502 Справка о стоимости ч_2
    # 0420502 Справка о стоимости ч_3
    # 0420502 Справка о стоимости ч_4
    # 0420502 Справка о стоимости ч_5
    # 0420502 Справка о стоимости ч_6
    # 0420502 Справка о стоимости ч_7
    # 0420502 Справка о стоимости ч_8
    # 0420502 Справка о стоимости ч_9
    # 0420502 Справка о стоимости _10
    # 0420502 Справка о стоимости _11
    # 0420502 Справка о стоимости _12

    period_begin = analiz_data_data(df_avancor.loc[18, 1])
    period_end = analiz_data_data(df_avancor.loc[18, 5])

    sheet_xbrl_period = wb_xbrl.sheetnames[8:18]

    for sheet in sheet_xbrl_period:
        ws_period = wb_xbrl[sheet]
        cell_period = df_matrica.loc[sheet, 'cell_period']
        ws_period[cell_period].value = period_begin + ', ' + period_end


# %%
# Анализ возможных ошибок при копировании данных

def corrector_scha_13(id_fond, wb_xbrl, df_avancor, df_id):
    """ Копируем количество паев с нужной точностью """
    # форма:
    # 0420502 Справка о стоимости _13

    # Точность указания кол-ва паев
    df_id_ws = df_id['ПИФ']
    fix = df_id_ws[df_id_ws[0] == id_fond][2].values[0]

    # записываем в форму
    ws_pai = wb_xbrl['0420502 Справка о стоимости _13']
    today_row, today_col = coordinate('I144')
    previous_row, previous_col = coordinate('K144')

    today = toFixed(df_avancor.loc[today_row, today_col], digits=fix)
    previous = toFixed(df_avancor.loc[previous_row, previous_col], digits=fix)

    ws_pai.cell(10, 3).value = today
    ws_pai.cell(10, 4).value = previous


# %%

def corrector_scha_00(wb):
    """ Убираем лишние нули (".00"), которые могут появлиться в результате анализа данных """
    # формы:
    # 0420502 Справка о стоимости чис
    # 0420502 Справка о стоимости _14
    # 0420502 Справка о стоимости _15
    # 0420502 Справка о стоимости _53

    def correct(row, col):
        # разбиваем на две части и оставляем только первую часть (без нулей)
        ws.cell(row, col).value = str(ws.cell(row, col).value).split('.')[0]

    # Корректируем реквизиты фонда: "№ лицензии", "ОКПО"
    # названия вкладок
    sheet_rekvizits = '0420502 Справка о стоимости чис'
    ws = wb[sheet_rekvizits]
    correct(10, 4)  # "№ лицензии"
    correct(10, 5)  # "ОКПО"

    # Корректируем номер кредитной организации (поле:"Регистрационный номер кредитной организации")
    # названия вкладок
    sheets_NKO = ['0420502 Справка о стоимости _14', '0420502 Справка о стоимости _15']
    row_start = 11
    for sheet in sheets_NKO:
        ws = wb[sheet]
        for row in range(row_start, ws.max_row):
            correct(row, 6)

    # Корректируем ИНН
    sheet = '0420502 Справка о стоимости _53'
    ws = wb[sheet]
    row_begin = 11
    col = 9
    for row in range(row_begin, ws.max_row):
        correct(row, col)


# %%
def corrector_scha_51(wb, df_id):
    """ Вставляем подробное описание имущества """
    # форма:
    # 0420502 Справка о стоимости _51

    """ 
    Вставляем подробное описание имущества.
    Колонка(3): 'Сведения, позволяющие определенно установить имущество'
    """
    # Определяем рабочий лист в конечном файле
    wb_sheetname = '0420502 Справка о стоимости _51'
    wb_ws = wb[wb_sheetname]

    # Определяем рабочий лист в фале-Идентификаторы
    df_id_sheetname = 'Иное имущество'
    df_id_ws = df_id[df_id_sheetname]

    wb_row_begin = 11  # начальная строка в конечном файле
    wb_column_id = 2  # колонка с идентификаторами в конечном файле
    wb_column = wb_column_id + 1  # колонка с описанием имущества в конечном файле

    df_row_begin = 1  # начальная строка в файле-Идентификаторы
    df_column_id = 0  # колонка с идентификаторами в файле-Идентификаторы
    df_column = df_column_id + 2  # колонка с описанием имущества в файле-Идентификаторы
    df_max_row = df_id_ws.shape[0]  # количество строк в файле-Идентификаторы

    # Перебираем строки с идентификаторами в конечном файле
    for wb_row in range(wb_row_begin, wb_ws.max_row):
        wb_cell = wb_ws.cell(wb_row, wb_column_id)
        # перебираем ячейки в файле-Идентификатор
        for df_row in range(df_row_begin, df_max_row):
            df_cell = df_id_ws.loc[df_row, df_column_id]
            # Сравниваем значения и если совпадают,
            # то записываем в конечный файл описание имущества из файла-Идентификатор
            if wb_cell.value == df_cell:
                wb_ws.cell(wb_row, wb_column).value = df_id_ws.loc[df_row_begin, df_column]

    """ 
    Убираем значение из итоговой строки.
    Колонка(4): 'Иное имущество - Количество в составе активов, штук'
    """
    wb_ws.cell(wb_ws.max_row, 4).value = ''


# %%
def corrector_scha_56to57(wb_xbrl, df_matrica, df_avancor, ERRORS, errors_file):
    """Записываем в формы ФИО подписантов"""
    # Формы:
    # 0420502 Справка о стоимости _56
    # 0420502 Справка о стоимости _57

    # отбор форм подписантов
    sheet_2_fio = wb_xbrl.sheetnames[61:63]

    for form in sheet_2_fio:
        # находим, используя матрицу, раздел в файле Аванкор, соответствующий выбранной форме
        title_1_name = df_matrica.loc[form, 'sheet_1_title']

        # количество строк в таблице Аванкор
        index_max = df_avancor.shape[0]
        # Номер строки с названием раздела в файле Аванкор
        title_row = razdel_name_row(df_avancor, title_1_name, index_max, ERRORS, errors_file)

        # номер колонки ячейки с данными в таблице Аванкор
        data_col = df_matrica.loc[form, 'cell1_col']
        data_col = column_index_from_string(data_col)

        # ФИО подписанта
        fio = df_avancor.loc[title_row, data_col]

        # записываем в форму xbrl
        ws_fio = wb_xbrl[form]
        row_fio, col_fio = coordinate(df_matrica.loc[form, 'cell_fio'])
        ws_fio.cell(row_fio, col_fio).value = fio


# %%
def corrector_scha_56to57_fio_full(wb_xbrl, file_id, ERRORS, errors_file):
    """ Вставляем в ячейки с подписантами полное ФИО вместо сокращенного ФИО"""

    # Формы:
    # 0420502 Справка о стоимости _56
    # 0420502 Справка о стоимости _57

    def insert_fio_full(sheet_name, cell_with_fio, wb_xbrl):
        """ Вставляем полное ФИО в ячейку"""

        def searche_fio_full():
            """ Ищем полное ФИО в файле с Идентификаторами """
            fio_full = None
            # Удаляем лишние пробелы из ФИО (если пробелы есть)
            try:
                fio_short_bp = fio_short.replace(' ', '')
            except AttributeError:
                error = f'форма "{sheet_name}", ФИО подписанта "{fio_short}" - не заполнено '
                print(error)
                ERRORS.append(error)
                write_errors(ERRORS, errors_file)
                return fio_full

            except:
                error = f'"{sheet_name}" - Неизвестная ошибка'
                print(error)
                ERRORS.append(error)
                write_errors(ERRORS, errors_file)
                return fio_full

            # Удаляем лишние пробелы из ФИО (если пробелы есть)
            for n in range(len(ws_df_id['ФИО_кратко'])):
                ws_df_id['ФИО_кратко'][n] = ws_df_id['ФИО_кратко'][n].replace(' ', '')

            try:  # находим нужное ФИО
                fio = ws_df_id[ws_df_id['ФИО_кратко'] == fio_short_bp]
                # переиндексируем df, начиная с '0'
                fio.index = range(0, len(fio))
                fio_full = fio.loc[0, 'ФИО_полностью']
            except KeyError:
                error = f'"{sheet_name}", полное ФИО подписанта "{fio_short}" не найдено'
                print(error)
                ERRORS.append(error)
                write_errors(ERRORS, errors_file)
            except:
                error = f'"{sheet_name}", Неизвестная ошибка'
                print(error)
                ERRORS.append(error)
                write_errors(ERRORS, errors_file)
            return fio_full

        ws_xbrl = wb_xbrl[sheet_name]
        fio_short = ws_xbrl[cell_with_fio].value
        fio_full = searche_fio_full()
        # Если находим полное ФИО, то вставляем его в форму xbrl
        if fio_full:
            ws_xbrl[cell_with_fio].value = fio_full
            print(fio_full)

    # Загружаем нужную страницу файла с Идентификаторами
    ws_df_id = pd.read_excel(file_id, sheet_name='ФИО', header=0)

    sheet_name = '0420502 Справка о стоимости _56'
    cell_with_fio = 'B7'
    insert_fio_full(sheet_name, cell_with_fio, wb_xbrl)

    sheet_name = '0420502 Справка о стоимости _57'
    cell_with_fio = 'B8'
    insert_fio_full(sheet_name, cell_with_fio, wb_xbrl)


# %%
def corrector_scha_57(wb_xbrl, df_id, id_fond):
    """ Проставляем id Фонда и реквизиты СпецДепа в форму"""
    # форма:
    # 0420502 Справка о стоимости _57

    # Проставляем id Фонда и реквизиты СпецДепа в форму
    ws_xbrl = wb_xbrl['0420502 Справка о стоимости _57']
    ws_id = df_id['ПИФ']

    # находим строку в df 'Идентификаторы' с реквизитами СД
    requisites = ws_id[(ws_id[0] == id_fond)]
    # переиндексируем df, начиная с '0'
    requisites.index = range(0, len(requisites))
    row = 8

    # вставляем идентификатор фонда
    ws_xbrl.cell(row, 1).value = id_fond
    # вставляем реквизиты СД
    # номера колонок в df 'Идентификаторы' начинается с "0", а ws_xbrl с "1"
    for col in range(3, 6):
        ws_xbrl.cell(row, col).value = str(requisites.loc[0, col])


# %%


# %%
# %%
# %%
# %%
# %%
# %%

if __name__ == "__main__":
    # corrector_scha_51(wb, df_id)
    pass
