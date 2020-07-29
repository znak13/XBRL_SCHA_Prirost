from module.functions import coordinate
from openpyxl.utils.cell import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

global log


def find_parametr(ws, row_begin, col):
    """ Находим название столбца при определении возможной ошибки"""

    for row in range(1, 10):
        cell = ws.cell(row_begin - row, col).value

        # в пояснительной записке №3 вместо номера столбца укзано "Содержание"
        if str(cell).isdigit() or str(cell) == "Содержание":
            # название столбца находится на строчку выше
            row_param = row_begin - row - 1
            cell = ws.cell(row_param, col).value
            return cell
        # в пояснительной записке №2 нет номера столбца
        if str(cell).startswith('Сведения о событиях'):
            # название столбца находится в этой ячейке
            row_param = row_begin - row
            cell = ws.cell(row_param, col).value
            return cell


def check_errors(ws, cell_avancor, row_begin, r, c, col_begin):
    """ Проверка на наличие ошибок """

    # Наименование колонки в таблице xbrl
    col_name = find_parametr(ws, row_begin, col_begin + c)

    if (str(cell_avancor).startswith("Не установлен")
        or str(cell_avancor) == 'nan'
        or str(cell_avancor) == '-') \
            and col_name != 'Примечание':

        log.error(f'"{ws.title}"; '
                  f'строка({row_begin + r}), '
                  f'колонка({c + 1}) --> '
                  f'параметр: "{col_name}"'
                  f' ==> "{cell_avancor}"')


def empty_cell(ws, cellBegin, cellEnd):
    """ Проверяем является ли ячейка пустой"""

    rowBegin, colBegin = coordinate_to_tuple(cellBegin)
    rowEnd, colEnd = coordinate_to_tuple(cellEnd)

    drow = rowBegin + (rowEnd - rowBegin) + 1
    dcol = colBegin + (colEnd - colBegin) + 1

    for row in range(rowBegin, rowBegin + (rowEnd - rowBegin) + 1):
        for col in range(colBegin, colBegin + (colEnd - colBegin) + 1):
            cellData = ws.cell(row, col).value
            if not cellData:
                log.error(f'"{ws.title}" --> пустая ячейка "{get_column_letter(col) + str(row)}"')
