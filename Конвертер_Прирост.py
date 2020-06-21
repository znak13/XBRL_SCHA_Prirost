import pandas as pd
import openpyxl
import os
from openpyxl.styles import Alignment
import sys
from module.analiz_data import *
import module.functions as fun

from tkinter.filedialog import askopenfilename
from tkinter import Tk
# Tk().withdraw()
root = Tk()
root.withdraw()

# %%
reportDir = "./#Отчетность"
# Выбор файла, созданного Аванкор
print(f'Выбор файла, созданного в Аванкор...')
# show an "Open" dialog box and return the path to the selected file
file_avancor = askopenfilename(initialdir=reportDir,
                               title="Выбор файла, созданного в Аванкор...",
                               filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
# Имя файла без пути к нему
if file_avancor:
    print(f'...выбран файл: {os.path.basename(file_avancor)}')
else:
    print(f'...файл не выбран')
    sys.exit()

# Загружаем данные из файла Аванкор
df_avancor = pd.read_excel(file_avancor, sheet_name='TDSheet', header=None)
# устанавливаем начальный индекс не c '0', а c '1'
df_avancor.index += 1
df_avancor.columns += 1

row_start_av = 26
row_end_av = 97
col_2_av = 5
col_3_av = 6
data = {}
for row in range(row_start_av, row_end_av + 1):
    cell_2 = df_avancor.loc[row, col_2_av]
    cell_3 = df_avancor.loc[row, col_3_av]
    if str(cell_2) != 'nan' and cell_3 != 0:
        data[cell_2] = analiz_data_all(cell_3)

# %%
# Выбор файла таблицы xbrl
print(f'Выбор файла таблицы xbrl c ЗАГРУЖЕННЫМИ(!) данными СЧА')
# show an "Open" dialog box and return the path to the selected file
file_xbrl = askopenfilename(title="Выбор файла таблицы xbrl c ЗАГРУЖЕННЫМИ(!) данными СЧА...",
                            filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
# Имя файла без пути к нему
print(f'...выбран файл: {os.path.basename(file_xbrl)}')

# Загружаем данные из файла таблицы xbrl
# file_xbrl = '0420502_0420503_Квартал_ЗПИФ_Дон.xlsx'
wb_xbrl = openpyxl.load_workbook(filename=file_xbrl)
# заполняем только одну форму (!)

# 0420503 Отчет о приросте об у_4 - SR_0420503_R3
urlSheets = fun.codesSheets(wb_xbrl)
sheetName = fun.sheetNameFromUrl(urlSheets, 'SR_0420503_R3')

ws_xbrl = wb_xbrl['0420503 Отчет о приросте об у_4']
row_start_xbrl = 7
row_end_xbrl = 54
col_2_xbrl = 2
col_3_xbrl = 3

for row in range(row_start_xbrl, row_end_xbrl + 1):
    # номер строки в xbrl
    cell = ws_xbrl.cell(row, col_2_xbrl).value
    # сравниваем номера строк в Аванкоре и xbrl
    if cell in data.keys():
        ws_xbrl.cell(row, col_3_xbrl).value = data[cell]
        # Форматируем ячейку
        ws_xbrl.cell(row, col_3_xbrl).alignment = Alignment(horizontal='right')

# %%

# Сохраняем результаты в файл отчетности xbrl
try:
    # os.chdir (os.path.dirname (file_fond_name))
    wb_xbrl.save(file_xbrl)
    print('-------------------ГОТОВО!!!----------------------')

except PermissionError:
    print(f'{"=" * 100}\n'
          f'ОШИБКА ДОСТУПА К ФАЙЛУ\n'
          f'(файл открыт в другой программе - закройте файл!)\n'
          f'{"=" * 100} ')
    input()

# root.destroy()