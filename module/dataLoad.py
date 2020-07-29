import pandas as pd
import os
import sys
import openpyxl
import shutil
# from Конвертер_СЧА import file_open
from tkinter.filedialog import askopenfilename, asksaveasfilename


def load_file(folder=None,
              fileName=None,
              sheet_name=None,
              index_col=None,
              header=None):
    """ Закружаем данные из файла 'excel' """

    file = folder + '/' + fileName
    df = pd.read_excel(file, sheet_name=sheet_name, index_col=index_col, header=header)

    return df


def pathToFile(up=1, folder=None):
    """ Путь к файлу, расположенному в папке 'folder' """
    # up = 1  # кол-во каталогов "вверх"
    # folder = название папки: например - 'Шаблоны'

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
    # Имя файла без расширения == идентификатор фонда
    id_fond = idFromFileName(file_avancor)

    print(f'выбран файл: {file_avancor}')
    return file_open, id_fond


# ========================================================================
#%%
# def load_data():
#
#     # Путь к папке с 'Шаблоны' из папки 'module'
#     folderID = pathToFile(up=1, folder='Шаблоны')
#     # ---------------------------------------------------------
#     # Загрузка файла-Матрицы
#     df_matrica = load_file(folder=folderID,
#                            fileName='Матрица.xlsx',
#                            sheet_name='0420502',
#                            index_col=1,
#                            header=0)
#     # ---------------------------------------------------------
#     # Загрузка файла с Идентификаторами
#     df_identifier = load_file(folder=folderID,
#                               fileName='Идентификаторы.xlsx',
#                               sheet_name=None,
#                               index_col=None,
#                               header=None)
#     file_id = folderID + '/' + 'Идентификаторы.xlsx'
#     # ---------------------------------------------------------
#     # Выбор файла, созданного в Аванкор
#     file_avancor, id_fond = openAvancore(df_identifier)
#     # ---------------------------------------------------------
#     # Загрузка файла-Аванкор
#     path_to_file_avancor = os.path.dirname(file_avancor)
#     # отбрасываем путь к файлу
#     fileNameAvancor = os.path.basename(file_avancor)
#     df_avancor = load_file(folder=path_to_file_avancor,
#                            fileName=fileNameAvancor,
#                            sheet_name='TDSheet',
#                            index_col=None,
#                            header=None)
#     # устанавливаем начальный индекс не c 0, а c 1
#     df_avancor.index += 1
#     df_avancor.columns += 1
#     # ---------------------------------------------------------
#     # добавляем к названию файла ошибок идентификатор фонда
#     errors_file = os.path.splitext(file_avancor)[0] + " - " + 'errors.txt'
#     # ---------------------------------------------------------
#     # Выбираем имя создаваемого файла-отчетности
#     file_fond_name = newFileName()
#     # ---------------------------------------------------------
#     # Используем файл-Шаблон
#     file_shablon = '0420502_0420503_Квартал - 3_1.xlsx'
#     # file_shablon = '0420502_0420503_Квартал - 3_2.xlsx'
#     print(f'Используем шаблон: {file_shablon}')
#     # ---------------------------------------------------------
#     # Создаем новый файл отчетности 'file_fond_name',
#     # создав копию шаблона 'file_shablon'
#     shutil.copyfile(folderID + '/' + file_shablon, file_fond_name)
#     # ---------------------------------------------------------
#     # Загружаем данные из файла таблицы xbrl
#     wb = openpyxl.load_workbook(filename=file_fond_name)
#     # ---------------------------------------------------------
#     return id_fond, file_id, df_identifier, df_avancor, df_matrica, wb, file_fond_name, errors_file

#%%
def load_data():
    """ Выбор файлов и загрузка данных"""

    # ---------------------------------------------------------
    # Загрузка файла с Идентификаторами
    df_identifier = load_file(folder=folderShablon,
                              fileName=fileID,
                              sheet_name=None,
                              index_col=None,
                              header=None)
    file_id = folderShablon + '/' + fileID
    # ---------------------------------------------------------
    # Выбор файла, созданного в Аванкор
    file_avancor, id_fond = openAvancore(df_identifier)
    # ---------------------------------------------------------
    # Загрузка файла-Аванкор
    path_to_file_avancor = os.path.dirname(file_avancor)
    # отбрасываем путь к файлу
    fileAvancor = os.path.basename(file_avancor)
    # создаем df
    df_avancor = load_file(folder=path_to_file_avancor,
                           fileName=fileAvancor,
                           sheet_name=fileAvancore_sheetNname,
                           index_col=None,
                           header=None)
    # устанавливаем начальный индекс не c 0, а c 1
    df_avancor.index += 1
    df_avancor.columns += 1
    # ---------------------------------------------------------
    # путь к файлам-отчетности
    path_to_rerort = path_to_file_avancor + '/'
    # ---------------------------------------------------------
    # Выбираем имя создаваемого файла-отчетности
    file_fond_name = newFileName()
    # ---------------------------------------------------------
    # Создаем новый файл отчетности 'file_fond_name',
    # создав копию шаблона 'file_shablon'
    shutil.copyfile(folderShablon + '/' + fileShablon, file_fond_name)
    # ---------------------------------------------------------
    # Загружаем данные из файла таблицы xbrl
    wb = openpyxl.load_workbook(filename=file_fond_name)
    # ---------------------------------------------------------
    return id_fond, df_identifier, df_avancor, wb, file_fond_name, path_to_rerort


#%%
def idFromFileName(fileName):
    """ Поиск в названии файла идентификатор фонда"""
    # Имя файла без расширения
    id_fond = os.path.splitext(fileName)[0]
    # разбиваем название на части
    id_fond = id_fond.split('_')
    # соединяем первые две части: должен получиться идентификатор фонда
    id_fond = '_'.join(id_fond[:2])

    # Список всех идентификаторой фондов
    df_identifier = load_file(folder=folderShablon, fileName=fileID)
    all_id_fond = df_identifier['ПИФ'][0][1:].to_list()

    # Если в названии файла нет идентификатора, то прерываем программу
    if not (id_fond in all_id_fond):
        print(f'.......ERROR!.......')
        print(f'{"=" * 100}\n'              
              f'Файл не сформирован!\n'
              f'В названии файла: "{fileName}" неверно указан идентификатор фонда!\n'
              f'(проверьте название файла)\n'
              f'{"=" * 100}')
        sys.exit()


    return id_fond



# ============================================================================
# Переменные

# папка с Шаблонами и Идентификаторами
folderWithShablon = 'Шаблоны'
# Путь к папке с 'Шаблоны' из папки 'module'
folderShablon = pathToFile(up=1, folder=folderWithShablon)
# название файла с Шаблоном отчетности
fileShablon = '0420502_0420503_Квартал - 3_1.xlsx'
# fileShablon = '0420502_0420503_Квартал - 3_2.xlsx'
print(f'Используем шаблон: {fileShablon}')


# название файла с Идентификаторами
fileID = 'Идентификаторы.xlsx'
# вкладка в файле-отчетности, созданном в Аванкоре
fileAvancore_sheetNname = 'TDSheet'
# имя файла с ошибками
# errorsFile = 'errors.txt'
# ---------------------------------------------------------

if __name__ == '__main__':
    pass
