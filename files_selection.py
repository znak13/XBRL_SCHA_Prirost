import os
import pandas as pd

from tkinter import *
from tkinter.ttk import Combobox
from tkinter.filedialog import askopenfilename

from module.functions import fond_id_search
from module.globals import *


# # styles
# style = Style()
# style.configure("GRN.TLabel", background="#ACF059")
# style.configure("GRN.TFrame", background="#ACF059")
# style.configure("BLK.TFrame", background="#595959")

class ReportFiles(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.parent = parent
        self.fondID = ''
        self.file_new_name = ''
        self.dir_name = ''
        self.file_Avancore_scha = ''
        self.file_path_Avancore = ''
        self.file_Avancore_prirost = ''
        self.file_xbrl = ''
        self.file_path_xbrl = ''
        self.todo = False
        self.initUI()

    def initUI(self):

        self.parent.title("Выбор параметров отчета")
        self.pack(fill=BOTH, expand=True)

        # ------------------------------------------------
        # Идентификатор Фонда:
        pifIDFrame = Frame(self, height=60, bg='')
        pifIDFrame.pack(side='top', fill='x')

        lbYear = Label(pifIDFrame, text="Идентификатор Фонда:", width=20, anchor=W)
        lbYear.pack(side=LEFT, padx=5, pady=5)

        self.combo = Combobox(pifIDFrame, width=25, foreground='red')
        # Загружаем список идентификаторов фонда из файла с идентификаторами
        # sheet_name = 'ПИФ'
        # df_pifID = pd.read_excel(dir_shablon + fileID, sheet_name=sheet_name, header=None)
        # pifID_old = list(df_pifID[0].to_list())[1:]
        pifID = fond_id_search()

        self.combo['values'] = pifID
        self.combo.pack(side=LEFT, padx=5)
        # вариант по умолчанию
        # self.combo.current(0)
        # self.fondID = self.combo.get()
        self.combo.bind("<<ComboboxSelected>>", self.enabl_fields)

        # ------------------------------------------------
        # Фрейм выбора файла-Аванкор-СЧА
        self.fileInputFrame("Выбрать файл-Аванкор-СЧА...",
                            "Файл-Аванкор-СЧА:", 1)
        # ------------------------------------------------
        # Фрейм выбора файла-Аванкор-ПРИРОСТ
        self.fileInputFrame("Выбрать файл-Аванкор-Прирост...",
                            "Файл-Аванкор-Прирост:", 2)
        # ------------------------------------------------
        # Фрейм выбора файла-xbrl с загруженными данными о СЧА
        self.fileInputFrame("Выбрать файл-xbrl с загруженными данными о СЧА...",
                            "Файл-xbrl:", 3)
        # ------------------------------------------------
        # Фрейм имени нового файла
        NewFileFrame = Frame(self, height=40, bg='')
        NewFileFrame.pack(side='top', fill='x')

        lbFile = Label(NewFileFrame, text="Имя нового файла:", width=16, anchor=W)
        lbFile.grid(column=0, row=0, sticky=W, padx=5)
        # поле ввода имени файла
        self.entry = Entry(NewFileFrame, width=32, state=DISABLED)
        self.entry.grid(column=1, row=0, sticky=NW, padx=5, pady=5)

        # ------------------------------------------------

        closeButton = Button(self, text="Close", height=1, width=10, command=self.doClose)
        closeButton.pack(side=RIGHT, anchor=S, padx=5, pady=5)
        okButton = Button(self, text="OK", height=1, width=10, command=self.doOK)
        okButton.pack(side=RIGHT, anchor=S, padx=0, pady=5)

    def enabl_fields(self, *args):
        self.entry['state'] = NORMAL

    def getEntry(self):
        """считываем имя нового файла"""
        self.file_new_name = self.entry.get()
        # Добавляем расширение файла
        if self.file_new_name:
            if self.file_new_name.endswith('.xlsx'):
                pass
            else:
                self.file_new_name += '.xlsx'

    # Кнопки выхода
    def doClose(self):
        self.todo = False
        print('Close')
        # self.parent.quit()
        self.parent.destroy()

    def doOK(self):
        self.todo = True
        self.fondID = self.combo.get()  # считываем fondID
        self.getEntry()  # считываем имя нового файла
        print('OK')
        # self.parent.quit()
        self.parent.destroy()

    def fileInputFrame(self, title, Buttontxt, fileType):

        def file_dir(fileType):
            """ имя выбранного файла и путь """
            fileName = askopenfilename(
                title=title,
                filetypes=(("xlsx files", "*.xlsx"), ("All files", "*.*")))
            if fileName:
                # путь к файлу
                self.dir_name = os.path.dirname(fileName)
                lbPath['text'] = self.dir_name
                # имя файла
                file_new = os.path.basename(fileName)
                lbFile['text'] = file_new

                if fileType == 1:
                    self.file_Avancore_scha = file_new
                    self.file_path_Avancore = self.dir_name
                elif fileType == 2:
                    self.file_Avancore_prirost = file_new
                    self.file_path_Avancore_prirost = self.dir_name
                elif fileType == 3:
                    self.file_xbrl = file_new
                    self.file_path_xbrl = self.dir_name

        # ------------------------------------------------
        # Кнопка выбора файла
        ButtonFrame = Frame(self, height=40, bg='')
        ButtonFrame.pack(side='top', fill='x')

        buttFile = Button(ButtonFrame, text=Buttontxt, width=20, anchor=W,
                          command=lambda: file_dir(fileType))
        buttFile.grid(column=0, row=0, sticky=W, padx=5, pady=5)

        # ------------------------------------------------
        # Метка выбора файла-Аванкор
        NamefileFrame = Frame(self, height=40, bg='')
        NamefileFrame.pack(side='top', fill='x')

        lbPathInfo = Label(NamefileFrame, text="имя файла:", width=10, anchor=W)
        lbPathInfo.grid(column=0, row=1, sticky=NW, padx=5)

        lbPathInfo = Label(NamefileFrame, text="путь к файлу:", width=10, anchor=W)
        lbPathInfo.grid(column=0, row=2, sticky=NW, padx=5, pady=5)

        lbFile = Label(NamefileFrame, text='', width=30, anchor=W)
        lbFile.grid(column=1, row=1, sticky=NW)

        lbPath = Label(NamefileFrame, text='', width=30, anchor=NW, height=3,
                       justify=LEFT, wraplength=200)
        lbPath.grid(column=1, row=2, rowspan=3, sticky=W, pady=5)


# ======================================================================
if __name__ == '__main__':
    pass
