import module.dataLoad as ld
import module.forms_maker._0420502_SCHA as scha
import module.forms_maker._0420502_Rasshifr as rf
import module.forms_maker._0420502_Podpisant as pp
import module.forms_maker._0420502_Zapiski as zap
import module.forms_maker._0420503_Prirost as prst

from module import logger
import builtins

from tkinter import Tk
root = Tk()
root.withdraw()
# ----------------------------------------------------------
# Загружвем данные
id_fond, df_identifier, df_avancor, wb, \
file_fond_name, path_to_rerort = ld.load_data()
# ----------------------------------------------------------

# Включаем логировние
log = logger.create_log(path=path_to_rerort,
                        file_log=id_fond+'_errors.log',
                        file_debug=id_fond+'_debug.log')
# устанавливаем 'log' как глобальную переменную (включая модули)
builtins.log = log

# Формируем итоговые формы СЧА
scha.scha(wb, id_fond, df_identifier, df_avancor)

# Формируем итоговые формы-расшифровки
rf.rashifr(wb, df_avancor, df_identifier, id_fond)

# Формирование форм СЧА - падписанты
pp.podpisant(wb, df_avancor, df_identifier, id_fond)

# Формирование форм СЧА - пояснительные записки
zap.zapiski_new(wb, df_avancor, id_fond)

# Удаляем формы-Прироста, которые не заполняются
# (остальные формы Прироста формируются отдельно)
prst.prirost(wb)

# Сохраняем результат
wb.save(file_fond_name)

