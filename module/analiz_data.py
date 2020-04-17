""" Анализ данных в ячейке """

from datetime import datetime


# ==================================================================================
def toFixed(number, digits=0):
    """Фиксируем количество знаков после запятой"""
    return f"{number:.{digits}f}"


# ==================================================================================

def analiz_data_data(cell):
    """ Конвертер даты """

    if type(cell) == str:
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
    if type(cell) == datetime:
        cell = cell.strftime("%Y-%m-%d")
        return cell

    return False


# ==================================================================================
def analiz_data_valuta(cell):
    """ Конвертер валюты и страны """

    if cell == 'RUB':
        cell = '643-RUB'
        return cell
    return False


# ==================================================================================
def analiz_data_strana(cell):
    """ Конвертер валюты и страны """

    if cell == 'RUS':
        cell = '643'
        return cell
    return False


# ============================================================
def analiz_data_number_point(cell):
    """ Конвертер десятичной части """

    if type(cell) == str:
        if len(cell.split(',')) == 2 or len(cell.split('.')) == 2:  # в тексте есть ОДНА запятая или точка
            cell = cell.replace(' ', '')  # удаляем пробелы (если есть)
            cell = cell.replace(',', '.', 1)  # заменяем запятую на точку
            try:  # пробуем преобразовать в число с 2-мя знаками после запятой
                return f'{float(cell):.2f}'
            except ValueError:
                # ValueError: could not convert string to float: '...строка...'
                return False
        return False
    return False


# Предыдущий вариант функции:
#
# def analiz_data_number_point(cell):
#     """ Конвертер десятичной части: меняем ',' на '.' """
#     # type (cell) == str
#
#     if type(cell) == str:
#         cell = cell.replace(' ', '')
#         # если в тексте разделяется знаком"," на две части
#         # если из первой части убрать один знак"-" (если "-" есть) и останутся одни цифры и
#         # если вторая часть - это одни цифры
#         if len(cell.split(',')) == 2 and \
#                 cell.split(',')[0].replace('-', '', 1).isdigit() and \
#                 cell.split(',')[1].isdigit():
#             part1 = cell.split(',')[0]
#             part2 = cell.split(',')[1]
#
#             if len(part2) == 1:  # если указан только один знак после запятой
#                 part2 = part2 + '0'
#             cell = part1 + '.' + part2
#
#             return cell.replace(' ', '')  # удаляем пробелы
#     return False


# ==================================================================================
def analiz_data_number_00(cell):
    """ Конвертер десятичной части
    фиксируем 2 знака после запятой
    # type (cell) == 'float' или  'int'  """

    if type(cell) == float or type(cell) == int:
        return str(toFixed(cell, 2))
    return False


# ==================================================================================
def analiz_data_number_shtuk(cell):
    """ Конвертер количество в штуках (удаляем побелы)"""

    if type(cell) == str:
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
    # print (analiz_data_all ('23 551 234') )

    # print ( analiz_data_number_point ('14534,56' ) )

    # q = analiz_data_all(23.866)
    # qq = str(q).split('.')[0]

    # from datetime import datetime, date, time
    # df_avancor.loc[349, 4].strftime("%Y-%m-%d")
    # dt.strftime("%A, %d. %B %Y %I:%M%p")

    pass