# import openpyxl
# from openpyxl.styles import Font

from module.dataCheck import check_errors
from module.dataCheck import red_error
from module.analiz_data import analiz_data_all
# from module.functions import coordinate, find_columns_numbers, \
#     razdel_name_row, start_data_row, end_data_row

from module.globals import *
global log


def copy_data(ws, df_avancor, rows_numbers, columns_numbers, row_begin, col_begin):
    """ Копирование данных из таблицы Аванкор в таблицу XBRL """

    for r, row in enumerate(rows_numbers):
        for c, col in enumerate(columns_numbers):
            cell_avancor = df_avancor.loc[row, col]
            check_errors(ws, cell_avancor, row_begin, r, c, col_begin)
            if cell_avancor != '-' and \
                    cell_avancor != "x" and \
                    cell_avancor != 0 and \
                    str(cell_avancor) != 'nan':
                ws.cell(row_begin + r, col_begin + c).value = \
                    analiz_data_all(df_avancor.loc[row, col])


# %%
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
    # info = 'Z= Идентификатор АИФ ПИФ-'
    if fondIDtxt in cell_id.value:
        cell_id.value = fondIDtxt + id_fond


if __name__ == "__main__":
    pass


