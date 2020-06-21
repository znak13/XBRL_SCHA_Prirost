import pandas as pd
import os
import openpyxl
import shutil
from module.functions import sys, write_errors
# from Конвертер_СЧА import file_open
from tkinter.filedialog import askopenfilename, asksaveasfilename


def load_file(folder='folder',
              fileName='fileName',
              sheet_name=None,
              index_col=None,
              header=None):
    """ Закружаем данные из файла"""

    file = folder + '/' + fileName
    df = pd.read_excel(file, sheet_name=sheet_name, index_col=index_col, header=header)

    return df


def pathToFile(up=1, folder='folder'):
    """ Путь к файлу от """
    # up = 1  # кол-во каталогов "вверх"
    # folder = 'Шаблоны'
    path_to_current_file = os.path.realpath(__file__)
    path_to_current_folder = os.path.dirname(path_to_current_file)
    path_to_folder = path_to_current_folder.split('\\')
    if up > 0:
        path_to_folder = path_to_folder[:-up]
    path_to_folder.append(folder)
    path_to_folder = '/'.join(path_to_folder)

    return path_to_folder


def newFileName():
    # название нового файла-отчетности xbrl
    print(f'Имя нового файла отчетности....: ', end='')
    file_fond_name = asksaveasfilename(title="Имя нового файла отчетности...",
                                       filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
    # Добавляем расширение файла
    lastPath = file_fond_name.split('/')[-1]
    if lastPath.split('.')[-1] != 'xlsx':
        file_fond_name = file_fond_name + '.xlsx'

    # Печатаем имя выбранного файла
    print(f'{os.path.basename(file_fond_name)}')

    return file_fond_name


def openAvancore(df_id):
    """ Выбор файла, созданного в Аванкор"""

    # Список всех идентификаторой фондов
    all_id_fond = df_id['ПИФ'][0][1:].to_list()

    # Выбираем файл, сформированный Аванкор
    print(f'Выбираем файл, сформированный Аванкор....'
          f'(файл должен начинаться с идентификатора фонда)')
    # show an "Open" dialog box and return the path to the selected file
    file_open = askopenfilename(initialdir="./#Отчетность",
                                title="Выбираем файл, сформированный Аванкор....",
                                filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
    # Имя файла без пути к нему
    file_avancor = os.path.basename(file_open)
    # Имя файла без расширения (идентификатор фонда)
    id_fond = os.path.splitext(file_avancor)[0]

    id_fond = id_fond.split('_', 2)
    id_fond = '_'.join(id_fond[:2])

    # Если в названии файла нет идентификатора, то прерываем программу
    if not (id_fond in all_id_fond):
        print('.......ERROR!.......')
        ERRORS.append(f'{"=" * 100}\n'
                      f'Файл не сформирован!\n'
                      f'В названии файла: "{file_avancor}" неверно указан идентификатор фонда!\n'
                      f'(проверьте название файла)\n'
                      f'{"=" * 100}')
        write_errors(ERRORS, errors_file)
        sys.exit("Ошибка в имени файла!")

    print(f'выбран файл: {file_avancor}')
    return file_open, id_fond


# ========================================================================
#%%
def load_data():

    # Путь к папке с 'Шаблоны' из папки 'module'
    folderID = pathToFile(up=1, folder='Шаблоны')
    # ---------------------------------------------------------
    # Загрузка файла-Матрицы
    df_matrica = load_file(folder=folderID,
                           fileName='Матрица.xlsx',
                           sheet_name='0420502',
                           index_col=1,
                           header=0)
    # ---------------------------------------------------------
    # Загрузка файла с Идентификаторами
    df_identifier = load_file(folder=folderID,
                              fileName='Идентификаторы.xlsx',
                              sheet_name=None,
                              index_col=None,
                              header=None)
    file_id = folderID + '/' + 'Идентификаторы.xlsx'
    # ---------------------------------------------------------
    # Выбор файла, созданного в Аванкор
    file_avancor, id_fond = openAvancore(df_identifier)
    # ---------------------------------------------------------
    # Загрузка файла-Аванкор
    path_to_file_avancor = os.path.dirname(file_avancor)
    # отбрасываем путь к файлу
    fileNameAvancor = os.path.basename(file_avancor)
    df_avancor = load_file(folder=path_to_file_avancor,
                           fileName=fileNameAvancor,
                           sheet_name='TDSheet',
                           index_col=None,
                           header=None)
    # устанавливаем начальный индекс не c 0, а c 1
    df_avancor.index += 1
    df_avancor.columns += 1
    # ---------------------------------------------------------
    # добавляем к названию файла ошибок идентификатор фонда
    errors_file = os.path.splitext(file_avancor)[0] + " - " + 'errors.txt'
    # ---------------------------------------------------------
    # Выбираем имя создаваемого файла-отчетности
    file_fond_name = newFileName()
    # ---------------------------------------------------------
    # Используем файл-Шаблон
    file_shablon = '0420502_0420503_Квартал - 3_1.xlsx'
    # file_shablon = '0420502_0420503_Квартал - 3_2.xlsx'
    print(f'Используем шаблон: {file_shablon}')
    # ---------------------------------------------------------
    # Создаем новый файл отчетности 'file_fond_name',
    # создав копию шаблона 'file_shablon'
    shutil.copyfile(folderID + '/' + file_shablon, file_fond_name)
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_fond_name)
    # ---------------------------------------------------------
    return id_fond, file_id, df_identifier, df_avancor, df_matrica, wb, file_fond_name, errors_file

#%%
def load_data_2():

    # Путь к папке с 'Шаблоны' из папки 'module'
    folderID = pathToFile(up=1, folder='Шаблоны')
    # ---------------------------------------------------------
    # Загрузка файла-Матрицы
    df_matrica = load_file(folder=folderID,
                           fileName='Матрица.xlsx',
                           sheet_name='0420502',
                           index_col='URL_end',
                           header=0)
    # ---------------------------------------------------------
    # Загрузка файла с Идентификаторами
    df_identifier = load_file(folder=folderID,
                              fileName='Идентификаторы.xlsx',
                              sheet_name=None,
                              index_col=None,
                              header=None)
    file_id = folderID + '/' + 'Идентификаторы.xlsx'
    # ---------------------------------------------------------
    # Выбор файла, созданного в Аванкор
    file_avancor, id_fond = openAvancore(df_identifier)
    # ---------------------------------------------------------
    # Загрузка файла-Аванкор
    path_to_file_avancor = os.path.dirname(file_avancor)
    # отбрасываем путь к файлу
    fileNameAvancor = os.path.basename(file_avancor)
    df_avancor = load_file(folder=path_to_file_avancor,
                           fileName=fileNameAvancor,
                           sheet_name='TDSheet',
                           index_col=None,
                           header=None)
    # устанавливаем начальный индекс не c 0, а c 1
    df_avancor.index += 1
    df_avancor.columns += 1
    # ---------------------------------------------------------
    # добавляем к названию файла ошибок идентификатор фонда
    errors_file = os.path.splitext(file_avancor)[0] + " - " + 'errors.txt'
    # ---------------------------------------------------------
    # Выбираем имя создаваемого файла-отчетности
    file_fond_name = newFileName()
    # ---------------------------------------------------------
    # Используем файл-Шаблон
    file_shablon = '0420502_0420503_Квартал - 3_1.xlsx'
    # file_shablon = '0420502_0420503_Квартал - 3_2.xlsx'
    print(f'Используем шаблон: {file_shablon}')
    # ---------------------------------------------------------
    # Создаем новый файл отчетности 'file_fond_name',
    # создав копию шаблона 'file_shablon'
    shutil.copyfile(folderID + '/' + file_shablon, file_fond_name)
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_fond_name)
    # ---------------------------------------------------------
    return id_fond, file_id, df_identifier, df_avancor, df_matrica, wb, file_fond_name, errors_file


#%%

# ============================================================================
# Глобальные переменные
# файл с ощибками
errors_file = 'errors.txt'
# Список ошибок
ERRORS = []
# ============================================================================

if __name__ == '__main__':
    id_fond, file_id, \
    df_id, df_avancor, df_matrica, wb, \
    file_fond_name, errors_file = load_data()
