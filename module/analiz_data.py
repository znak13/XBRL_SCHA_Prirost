""" Анализ данных в ячейке """

from datetime import datetime


# ==================================================================================
def toFixed(number, digits=0):
    """Фиксируем количество знаков после запятой"""
    return f"{number:.{digits}f}"


# ==================================================================================
def analiz_data_data(cell):
    """ Конвертер даты """

    # если данные в ячейке, в формате "str": '17.07.2020'
    if isinstance(cell, str):
        if len(cell.split('.')) == 3 and \
                cell.split('.')[0].isdigit() and \
                cell.split('.')[1].isdigit() and \
                len(cell.split('.')[1]) == 2 and \
                cell.split('.')[2].isdigit():
            part1 = cell.split('.')[0]
            part2 = cell.split('.')[1]
            part3 = cell.split('.')[2]
            part1, part3 = part3, part1
            cell = part1 + '-' + part2 + '-' + part3
            return cell

    # если данные в ячейке, в формате "дата": 2222-07-31 00:00:00
    if isinstance(cell, datetime):
        return cell.strftime("%Y-%m-%d")

    return False


# ==================================================================================
def analiz_data_valuta(cell):
    """ Конвертер валюты и страны """

    if cell == 'RUB':
        return '643-RUB'
    return False


# ==================================================================================
def analiz_data_strana(cell):
    """ Конвертер валюты и страны """

    if cell == 'RUS':
        return '643'
    return False


# ============================================================
def analiz_data_number_point(cell):
    """ Конвертер десятичной части """

    if isinstance(cell, str):
        # в тексте есть ОДНА запятая или точка: '1 234,567' или '1 234.567'
        if len(cell.split(',')) == 2 or len(cell.split('.')) == 2:
            cell = cell.replace(' ', '')  # удаляем пробелы (если есть): '1 234,567' или '1 234.567'
            cell = cell.replace(',', '.', 1)  # заменяем запятую на точку: '1 234,567' или '1 234.567'
            try:  # пробуем преобразовать в число с 2-мя знаками после запятой
                return f'{float(cell):.2f}'
            except ValueError:  # ValueError: could not convert string to float: '...строка...'
                return False
        return False
    return False

# ==================================================================================
def analiz_data_number_00(cell):
    """ Конвертер десятичной части
    фиксируем 2 знака после запятой
    # type (cell) == 'float' или  'int'  """

    if isinstance(cell, (float, int)):  # 123 или 123.4567
        return str(toFixed(cell, 2))
    return False


# ==================================================================================
def analiz_data_number_shtuk(cell):
    """ Конвертер количество в штуках (удаляем побелы)"""

    if isinstance(cell, str):  # '1 234'
        cell = cell.replace(' ', '')
        if cell.isdigit():  # если после удаления пробелов остаются только цифры
            return cell
    return False


# ==================================================================================

def analiz_data_all(cell):
    """ Перебор всех возможных анализов данных """
    cell_new = analiz_data_data(cell)
    if not cell_new:
        cell_new = analiz_data_valuta(cell)
        if not cell_new:
            cell_new = analiz_data_strana(cell)
            if not cell_new:
                cell_new = analiz_data_number_point(cell)
                if not cell_new:
                    cell_new = analiz_data_number_00(cell)
                    if not cell_new:
                        cell_new = analiz_data_number_shtuk(cell)
                        if not cell_new:
                            cell_new = cell
    return cell_new


# ==================================================================================

if __name__ == "__main__":
    pass
