import module.dataLoad as ld
import module.forms_maker.SCA_0420502 as scha
import module.forms_maker.Rasshifr_0420502 as rf
import module.forms_maker.Podpisant_0420502 as pp
import module.forms_maker.Zapiski_0420502 as zap
import module.forms_maker.Prirost_0420503 as prst

from tkinter import Tk
root = Tk()
root.withdraw()
# ----------------------------------------------------------
# Загружвем данные
id_fond, file_id, df_identifier, df_avancor, df_matrica, wb, \
file_fond_name, errors_file = ld.load_data_2()
# ----------------------------------------------------------

# Глобальная переменная (собираем ошибки)
ERRORS = []

# Формируем итоговые формы СЧА
scha.scha_itogi(id_fond, file_id, df_identifier, df_avancor, df_matrica, wb,
               file_fond_name, errors_file, ERRORS)

# Формируем итоговые формы-расшифровки
rf.scha_rashifr(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS)

# Формирование форм СЧА - падписанты
pp.podpisant(wb, df_matrica, df_avancor, df_identifier, id_fond, ERRORS, errors_file)

# Формирование форм СЧА - пояснительные записки
zap.zapiski(wb, df_matrica, df_avancor, ERRORS)

# Удаляем формы из Прироста, которые не заполняются
prst.prirost(wb)

# Сохраняем результат
wb.save(file_fond_name)

