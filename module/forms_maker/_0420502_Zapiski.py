""" Формирование форм СЧА - пояснительные записки"""
import module.functions as fun
import module.dataCopy as dcop
from openpyxl.utils.cell import coordinate_to_tuple  # 'D2' -> (2,4)

global log
# Формы:
# 0420502 Пояснительная записка к - SR_0420502_PZ_inf_fakt_sversh_oshib
# 0420502 Пояснительная записка_2 - SR_0420502_PZ_sved_sobyt_okaz_susshestv_vliayn_scha
# 0420502 Пояснительная записка_3 - SR_0420502_PZ_inaya_inf
# 0420502 Пояснительная записка_4 - SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch
# 0420502 Пояснительная записка_5 - SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch_2
# 0420502 Пояснительная записка_6 - SR_0420502_PZ_inf_fakt_raznoglas_so_spec_dep

# ==================================================================================
def zapiski_new(wb, df_avancor, id_fond):
    # **********************************************************************************
    def copyFromAvancore(ws, avancoreTitle, max_number, cell_start_row, cell_start_col):

        # кол-во строк и столбцов в файле Аванкор
        index_max = df_avancor.shape[0]
        collumn_max = df_avancor.shape[1]

        zapiski_null = []  # список пустых форм
        # for url, form in urlSheets.items():

        if avancoreTitle:
            # если ячейка в столбце "sheet_1_title" пустая, то "title_1_name == nan"
            # в этом случае: bool(title_1_name == title_1_name) == Fals
            # (...странно, но работает)

            # Номер строки с названием раздела в файле Аванкор"""
            title_row = fun.razdel_name_row(df_avancor, avancoreTitle, index_max)

            # находим номер первой строку с данными в файле Аванкор
            data_row = fun.start_data_row(df_avancor, index_max, title_row)

            # находим номер последней строки с данными в таблице Аванкор"""
            row_end = fun.end_data_row(df_avancor, index_max, data_row)

            # Если ячейка - (столбец:3, строка: над последней строкой)
            # не равна "2" (это заголовок второй колонки),
            # то значит в таблице есть данные для копирования
            # if df_avancor.loc[row_end - 1, 3] != '2':

            # Если ячейка - (столбец:2, строка: над последней строкой)
            # не равна "1" (это заголовок 1-ой колонки таблицы),
            # или текст в ячейки начинается с 'Оценочная' (предпоследняя таблица),
            # то значит в таблице есть данные для копирования
            cell = str (df_avancor.loc[row_end - 1 , 2])
            if cell != '1' and \
                    not cell.startswith('Оценочная'):

                # список всех номеров строк в таблице Аванкор
                rows_numbers = [x for x in range(data_row, row_end)]

                # Номера колонок в таблице Аванкор, за исключением пустых
                columns_numbers = fun.find_columns_numbers(df_avancor, collumn_max, max_number,
                                                           data_row,
                                                           data_col=3)

                # Копирование данных из таблицы Аванков в таблицу XBRL
                dcop.copy_data(ws, df_avancor, rows_numbers, columns_numbers, cell_start_row,
                               cell_start_col)

                log.error(f'"{ws.title}" --> отсутствует "Идентификатор строки"')

            else: # в файле Аванкор Раздел пустой
                return False
        else: # в файле Аванкор Раздела отсутствует
            return False

        return True

    def makeForm(shortURL, avancoreTitle, max_number, cell_start):
        cell_start_row, cell_start_col = coordinate_to_tuple(cell_start)
        sheetName = fun.sheetNameFromUrl(urlSheets, shortURL)  # имя вкладки
        ws = wb[sheetName]
        print(f'{sheetName} - {shortURL}')

        # Копируем данные из файла Аванкор
        if copyFromAvancore(ws, avancoreTitle, max_number, cell_start_row, cell_start_col):
            # Записываем в форму идентификатор фонда
            dcop.copy_id_fond_to_tbl(ws, id_fond)

        else:
            # Если ничего скопировано не было, то Раздел пуст.
            # Удаляем вкладку
            wb.remove(ws)

    # **********************************************************************************
    def zapiska_01():
        """0420502 Пояснительная записка к - SR_0420502_PZ_inf_fakt_sversh_oshib"""

        shortURL = 'SR_0420502_PZ_inf_fakt_sversh_oshib'  # код вкладки
        avancoreTitle = 'Информация о фактах совершения ошибок, потребовавших ' \
                        'перерасчета стоимости чистых активов, а также о принятых ' \
                        'мерах по исправлению и последствиях исправления таких ошибок'
        # цифра в последней колонке таблицs XBRL
        # (количество колонок для копирования)
        max_number = 3
        # координаты первой ячейки в таблице XBRL
        cell_start = "B9"

        makeForm(shortURL, avancoreTitle, max_number, cell_start)

    # **********************************************************************************
    def zapiska_02():
        """0420502 Пояснительная записка_2 - SR_0420502_PZ_sved_sobyt_okaz_susshestv_vliayn_scha"""
        shortURL = 'SR_0420502_PZ_sved_sobyt_okaz_susshestv_vliayn_scha'  # код вкладки
        avancoreTitle = 'Сведения о событиях, которые оказали существенное влияние ' \
                        'на стоимость чистых активов акционерного инвестиционного фонда ' \
                        'или чистых активов паевого инвестиционного фонда'

        # цифра в последней колонке таблицs XBRL
        # (количество колонок для копирования)
        max_number = 1
        # координаты первой ячейки в таблице XBRL
        cell_start = "B8"

        makeForm(shortURL, avancoreTitle, max_number, cell_start)

    # **********************************************************************************
    def zapiska_03():
        """0420502 Пояснительная записка_3 - SR_0420502_PZ_inaya_inf"""
        shortURL = 'SR_0420502_PZ_inaya_inf'  # код вкладки
        avancoreTitle = 'Иная информация'

        # цифра в последней колонке таблицs XBRL
        # (количество колонок для копирования)
        max_number = 1
        # координаты первой ячейки в таблице XBRL
        cell_start = "B8"

        makeForm(shortURL, avancoreTitle, max_number, cell_start)

    # **********************************************************************************
    def zapiska_04():
        """0420502 Пояснительная записка_4 - SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch"""
        shortURL = 'SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch'  # код вкладки
        avancoreTitle = 'Информация о требованиях и обязательствах по опционным ' \
                        'и (или) фьючерсным договорам (контрактам)'

        # цифра в последней колонке таблицs XBRL
        # (количество колонок для копирования)
        max_number = 10
        # координаты первой ячейки в таблице XBRL
        cell_start = "B10"

        makeForm(shortURL, avancoreTitle, max_number, cell_start)

    # **********************************************************************************
    def zapiska_05():
        """0420502 Пояснительная записка_5 - SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch_2"""
        shortURL = 'SR_0420502_PZ_inf_treb_i_obyaz_opz_fiuch_2'  # код вкладки
        avancoreTitle = False

        # цифра в последней колонке таблицs XBRL
        # (количество колонок для копирования)
        max_number = 1
        # координаты первой ячейки в таблице XBRL
        cell_start = "B7"

        makeForm(shortURL, avancoreTitle, max_number, cell_start)

    # **********************************************************************************
    def zapiska_06():
        """0420502 Пояснительная записка_6 - SR_0420502_PZ_inf_fakt_raznoglas_so_spec_dep"""
        shortURL = 'SR_0420502_PZ_inf_fakt_raznoglas_so_spec_dep'  # код вкладки
        avancoreTitle = 'Информация о фактах возникновения разногласий со ' \
                        'специализированным депозитарием при расчете стоимости ' \
                        'чистых активов, а также о принятых мерах по преодолению этих разногласий'

        # цифра в последней колонке таблицs XBRL
        # (количество колонок для копирования)
        max_number = 2
        # координаты первой ячейки в таблице XBRL
        cell_start = "B9"

        makeForm(shortURL, avancoreTitle, max_number, cell_start)

    # **********************************************************************************
    urlSheets = fun.codesSheets(wb)  # словарь - "код вкладки":"имя вкладки"
    zapiska_01()
    zapiska_02()
    zapiska_03()
    zapiska_04()
    zapiska_05()
    zapiska_06()


if __name__ == "__main__":
    pass
