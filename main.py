from tkinter import *
from files_selection import ReportFiles

from module import logger
import builtins

from module.globals import *

global log


# from loguru import logger
# logger.add("file_{time}.log")

# @ logger.catch()
def main():
    # Выполняем пока не сделан корректный выбор периода
    choice = False
    while not choice:

        root_files = Tk()
        root_files.geometry("350x500+600+300")
        file_set = ReportFiles(root_files)
        root_files.mainloop()

        # проверяем как закрыто окно и выбран ли файл
        # если файл не выбран, то повторяем цикл
        key = {'s': file_set.file_Avancore_scha,  # сча
               'p': file_set.file_Avancore_prirost,  # прирост
               'x': file_set.file_xbrl,  # xbrl-файл с сча
               'n': file_set.file_new_name}  # новый xbrl-файл

        # ----------------------------------------------------------
        # Включаем логировние
        log = logger.create_log(path=file_set.dir_name + '/',
                                file_log=file_set.file_new_name + '_' + file_set.fondID + log_endName_scha,
                                file_debug=file_set.file_new_name + '_' + file_set.fondID + debug_endName_scha)
        # устанавливаем 'log' как глобальную переменную (включая модули)
        builtins.log = log
        # ----------------------------------------------------------

        if not file_set.fondID and file_set.todo:
            print(f'Не выбран идентификатор фонда!\n'
                  f'Попробуйте еще раз.')
            continue

        try:
            if file_set.todo:
                print(file_set.fondID)
                # дописываем к названию файла идентификатор Фонда
                file_set.file_new_name = file_set.file_new_name + \
                                         '_' + file_set.fondID + \
                                         '.xlsx'

                if (key['s'] and key['p'] and key['n'] and not key['x']):
                    choice = True  # выбор сделан
                    print(f'Формируем отчет полностью:')
                    report = 'сча/прирост'
                    import Конвертер_СЧА_v2
                    import Конвертер_Прирост_v2
                    Конвертер_СЧА_v2.main(file_set.fondID,
                                          file_set.dir_name,
                                          file_set.file_new_name,
                                          file_set.file_Avancore_scha)
                    Конвертер_Прирост_v2.main(file_set.fondID,
                                              file_set.dir_name,
                                              file_set.file_Avancore_prirost,
                                              file_set.file_new_name,
                                              report)

                elif key['s'] and key['n'] and not key['p'] and not key['x']:
                    choice = True  # выбор сделан
                    print(f'Формируем только СЧА:')
                    report = 'сча'
                    import Конвертер_СЧА_v2
                    Конвертер_СЧА_v2.main(file_set.fondID,
                                          file_set.dir_name,
                                          file_set.file_new_name,
                                          file_set.file_Avancore_scha)

                elif key['p'] and key['x'] and not key['s'] and not key['n']:
                    choice = True  # выбор сделан
                    print(f'Формируем только Прирост:')
                    report = 'прирост'
                    import Конвертер_Прирост_v2
                    Конвертер_Прирост_v2.main(file_set.fondID,
                                              file_set.dir_name,
                                              file_set.file_Avancore_prirost,
                                              file_set.file_xbrl,
                                              report)

                else:
                    print(f'Файлы выбраны не правильно!\n'
                          f'Попробуйте еще раз.')
                    continue
            else:
                log.warning(f'Окно выбора закрыто: нажата кнопка "Close"')
                sys.exit()

        except AttributeError as e:  # окно выбора закрыто не кнопкой
            log.error(f'Окно выбора закрыто: (AttributeError) - {e}')
            # sys.exit()
        except Exception as e:
            log.error(f'Ошибка: {e}')
            # sys.exit()


if __name__ == '__main__':
    main()
