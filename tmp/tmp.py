import openpyxl
import sys

from openpyxl.utils.cell import coordinate_from_string  # ‘B12’ -> (‘B’, 12)
from openpyxl.utils import column_index_from_string  # 'B' -> 2


# %%

def fond_id_search(file_name='пиф_id.xlsx',
                   sheet_with_tables='списки',
                   tbl_name='pif'):
    """Список идентификаторов фонда"""

    wb = openpyxl.load_workbook(file_name)
    ws = wb[sheet_with_tables]
    # tt = ws._tables
    # for tbl in ws._tables:
    #     print(tbl)
    #     print(" : " + tbl.displayName)
    #     print("   -  name = " + tbl.name)
    #     print("   -  type = " + (tbl.tableType if isinstance(tbl.tableType, str) else 'n/a'))
    #     print("   - range = " + tbl.ref)
    #     print("   - #cols = %d" % len(tbl.tableColumns))
    #     for col in tbl.tableColumns:
    #         print("     : " + col.name)
    table_found = None
    try:
        for tbl in ws._tables:
            if tbl.name == tbl_name:
                table_found = tbl
                break
        if not table_found:
            print(f'Таблица "{tbl_name}" не найдена.')
            print(f'Невозможно построить список идентификаторов фондов.')
            raise SystemExit
    except SystemExit:
        sys.exit()

    # tbl_name = fond_id_tbl.tableColumns[0].name
    table_found_size = table_found.ref  # 'C19:D27'
    # start_cell = fond_tbl_ref.split(':')[0] # 'C19'

    cell_start = coordinate_from_string(table_found_size.split(':')[0])
    cell_end = coordinate_from_string(table_found_size.split(':')[1])
    col_id = column_index_from_string(cell_start[0])
    row_start = cell_start[1]
    row_end = cell_end[1]

    fond_id = []
    for row in range(row_start + 1, row_end + 1):
        fond_id.append(ws.cell(row, col_id).value)

    return fond_id


fond_id = fond_id_search()
