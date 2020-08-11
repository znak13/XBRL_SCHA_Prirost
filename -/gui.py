import tkinter as tk
import wx


def tkinter_fun():
    root = tk.Tk()
    # root.withdraw()  # hide the window

    e = tk.Entry(width=120)
    b = tk.Button(text="Преобразовать")
    l = tk.Label(bg='black', fg='white', width=20)

    def strToSortlist(event):
        s = e.get()
        s = s.split()
        s.sort()
        l['text'] = ' '.join(s)

    b.bind('<Button-1>', strToSortlist)

    e.pack()
    b.pack()
    l.pack()
    root.mainloop()

def wx_fun():

    APP_EXIT = 1
    VIEW_STATUS = 2
    VIEW_RGB = 3
    VIEW_SRGB = 4

    class Myframe(wx.Frame):
        def __init__(self, parent, title):
            super().__init__(parent, title=title)

            menuBar = wx.MenuBar()  # панель с меню

            # Меню "Файл"
            fileMenu = wx.Menu()
            # Добавление пунктов меню в меню "Файл"
            fileMenu.Append(wx.ID_NEW, "Новый")
            fileMenu.Append(wx.ID_OPEN, "Открыть")
            fileMenu.Append(wx.ID_SAVE, "Сохранить")
            # Добавление пункта меню с "подменю"
            expMenu = wx.Menu()
            expMenu.Append(wx.ID_ANY, "Экспорт изображения")
            expMenu.Append(wx.ID_ANY, "Экспорт видео")
            expMenu.Append(wx.ID_ANY, "Экспорт данных")
            fileMenu.AppendSubMenu(expMenu, "Экспорт")
            # Добавляем сепаратор (разделительная линия)
            fileMenu.AppendSeparator()

            # item = wx.MenuItem(fileMenu, wx.ID_EXIT, "Выход\tCtrl+Q", "Выход изпрограммы")
            # item.SetBitmap(wx.Bitmap('picture.png'))
            # fileMenu.Append(item)
            item = fileMenu.Append(APP_EXIT, "Выход\tCtrl+Q", "Выход изпрограммы")

            # Меню "Вид"
            viewMenu = wx.Menu()
            self.vStatus = viewMenu.Append(VIEW_STATUS, "Статус...", kind=wx.ITEM_CHECK)
            self.vRgb = viewMenu.Append(VIEW_RGB, "Тип_RGB", kind=wx.ITEM_RADIO)
            self.vSrgb = viewMenu.Append(VIEW_SRGB, "Тип_sRGB", kind=wx.ITEM_RADIO)

            # Добавление пунктов меню в "menuBar"
            menuBar.Append(fileMenu, "&File")
            menuBar.Append(viewMenu, "&Вид")

            # Установка панели "menuBar"
            self.SetMenuBar(menuBar)

            # Обработка событий
            self.Bind(wx.EVT_MENU, self.onStatus, id=VIEW_STATUS)
            self.Bind(wx.EVT_MENU, self.onImageType, id=VIEW_RGB)
            self.Bind(wx.EVT_MENU, self.onImageType, id=VIEW_SRGB)
            self.Bind(wx.EVT_MENU, self.onQuit, id=APP_EXIT)

        def onStatus(self, event):
            if self.vStatus.IsChecked():
                print("Выбран статус...")
            else:
                print("Статус не выбран")

        def onImageType(self, event):
            if self.vRgb.IsChecked():
                print("Выбран RGB")
            elif self.vSrgb.IsChecked():
                print("Выбран sRGB")

        def onQuit(self, event):
            self.Close()

    app = wx.App()
    frame = Myframe(None, "Hello World!")
    frame.Show()
    app.MainLoop()

# tkinter_fun()
wx_fun()