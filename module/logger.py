import logging


def create_log(path="",
               file_log='errors.log', mode='a',
               file_debug='debug.log', filemode='a'):
    """ Включаем логировние"""
    # path - путь к лог-файлам

    # logging.basicConfig(level=logging.DEBUG,
    #                     filename= 'log.txt',
    #                     filemode='w',
    #                     format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    logging.basicConfig(level=logging.DEBUG,
                        filename=path+file_debug,
                        filemode=filemode,
                        format='%(asctime)s - %(levelname)s - %(name)s - '
                               'модуль:"%(module)s":%(lineno)d - %(message)s')

    # Create a custom logger
    log = logging.getLogger('aya')

    # Create handlers (обработчики)
    c_handler = logging.StreamHandler() # в терминал
    f_handler = logging.FileHandler(path + file_log, mode=mode) # в файл
    c_handler.setLevel(logging.WARNING)
    f_handler.setLevel(logging.ERROR)

    # Create formatters and add it to handlers
    c_format = logging.Formatter('%(levelname)s - модуль:"%(module)s":%(lineno)d - %(message)s')
    f_format = logging.Formatter('%(levelname)s - %(message)s')
    c_handler.setFormatter(c_format)
    f_handler.setFormatter(f_format)

    # Add handlers to the logger
    log.addHandler(c_handler)
    log.addHandler(f_handler)

    # отключаем сообщения от логгера 'comtypes'
    logging.getLogger('comtypes').setLevel('CRITICAL')

    return log


# ===============================================================
if __name__ == "__main__":
    pass

    logging.basicConfig(level=logging.DEBUG,
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    # ---------------------------
    # список логеров
    for key in logging.Logger.manager.loggerDict:
        print(key)
    # ---------------------------
    # логирование исключения
    a = 5
    b = 0
    try:
        c = a / b
    except Exception as e:
        logging.exception("Exception occurred - ошибка")
    # ---------------------------
