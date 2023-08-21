# -*- coding: utf-8 -*-
import os

# список БД, который должен быть
list_db = [f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_19.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_20.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_21.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_22.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_23.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_24.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_25.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_PAUT_26.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_19.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_20.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_21.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_22.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_23.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_24.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_25.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OF_UTT_26.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_19.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_20.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_21.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_22.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_23.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_24.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_25.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_PAUT_26.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_19.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_20.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_21.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_22.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_23.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_24.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_25.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_ON_UTT_26.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_19.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_20.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_21.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_22.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_23.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_24.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_25.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_PAUT_26.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_19.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_20.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_21.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_22.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_23.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_24.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_25.sqlite',
           f'{os.path.abspath(os.getcwd())}\\DB\\reports_db_OS_UTT_26.sqlite']

# кортеж названий столбцов для поиска
tuple_utt_name_column_for_search = (
    'Line', 'Drawing', 'Item_description', 'Nominal_thickness', 'Diameter', 'North', 'South', 'West', 'East',
    'Top', 'Bottom', 'Extrados', 'Intrados', 'Column', 'Row', 'Center', 'Location', 'Section', 'Spot', 'S_NO',
    'Material', 'Date', 'P_ID', 'Comments', 'Remark', 'Size', 'Vertical', 'Horizontal', 'Shell', 'Plate',
    'Right', 'Left', 'Distance'
)

# список названий столбцов для UTT
utt_name_column = {'Line': ['LINE', 'TAG', 'CONTR', 'OBJ'],
                   'Drawing': ['DRAW', 'ISOM'],
                   'Item_description': ['ITEM', 'DESCR'],
                   'Diameter': ['ETER', 'INCH', 'ДИАМЕ', ],
                   'Nominal_thickness': ['NOM', 'THICK'],
                   'North': ['NOR'],
                   'South': ['SOU'],
                   'West': ['WES'],
                   'East': ['EAS'],
                   'Top': ['TOP'],
                   'Bottom': ['BOT'],
                   'Center': ['CENT'],
                   'Extrados': ['EXTR'],
                   'Intrados': ['INTR'],
                   'Spot': ['SPOT', 'CHECK'],
                   'Location': ['LOCA'],
                   'Section': ['SECT'],
                   'S_NO': ['S/NO', '№', 'NO'],
                   'Material': ['MATER'],
                   'Column': ['COL'],
                   'Row': ['ROW'],
                   'Date': ['DAT'],
                   'P_ID': ['P&ID'],
                   'Comments': ['COM'],
                   'Remark': ['REM'],
                   'Size': ['SIZE'],
                   'Vertical': ['VERT'],
                   'Horizontal': ['HORIZ'],
                   'Shell': ['SHEL'],
                   'Plate': ['PLAT'],
                   'Right': ['RIGHT'],
                   'Left': ['LEFT'],
                   'Distance': ['DIST']
                   }

# кортеж названий столбцов для поиска
tuple_paut_name_column_for_search = (
    'Location', 'Line', 'Drawing', 'Scanned_area', 'Item_description', 'Diameter', 'Nominal_thickness', 'Date',
    'Minimum_thickness', 'Maximum_thickness',
    'Average_thickness', 'Start_X', 'End_X', 'Start_Y', 'End_Y'
)

# список названий столбцов для PAUT
paut_name_column = {'Location': ['LOC'],
                    'Line': ['LINE', 'TAG', 'CONTR', 'OBJ'],
                    'Drawing': ['DRAW', 'ISOM'],
                    'Scanned_area': ['SCAN', 'AREA'],
                    'Item_description': ['ITEM', 'DESCR'],
                    'Diameter': ['ETER', 'INCH', 'ДИАМЕ', ],
                    'Nominal_thickness': ['NOM'],
                    'Date': ['DAT'],
                    'Minimum_thickness': ['MINI'],
                    'Maximum_thickness': ['MAXI'],
                    'Start_X': ['START X'],
                    'End_X': ['END X'],
                    'Start_Y': ['START Y'],
                    'End_Y': ['END Y'],
                    'Average_thickness': ['AVER']
                    }

# названия месяцев для дат репортов
month_alpha = {'JANUARY': 'JAN',
               'FEBRUARY': 'FEB',
               'MARCH': 'MAR',
               'APRIL': 'APR',
               'MAY': 'MAY',
               'JUNE': 'JUN',
               'JULY': 'JUL',
               'AUGUST': 'AUG',
               'SEPTEMBER': 'SEP',
               'OCTOBER': 'OCT',
               'NOVEMBER': 'NOV',
               'DECEMBER': 'DEC'}

# числовое представление месяцев
month_numeric = {'JANUARY': '01',
                 'FEBRUARY': '02',
                 'MARCH': '03',
                 'APRIL': '04',
                 'MAY': '05',
                 'JUNE': '06',
                 'JULY': '07',
                 'AUGUST': '08',
                 'SEPTEMBER': '09',
                 'OCTOBER': '10',
                 'NOVEMBER': '11',
                 'DECEMBER': '12'}

faq_text = '- Для чего эта программа?\n' \
           'Для быстрого поиска измеренных значений и чертежей UTT и PAUT начиная с 2019 года.\n\n' \
           '- Как искать данные?\n' \
           'Заполнить одно или несколько полей для поиска. Для самого быстрого поиска рекомендуется всегда заполнять ' \
           'поле "Юнит"!\n' \
           'Нет необходимости указывать полное название номеров линий, чертежей и т.д. Достаточно указать последние ' \
           'несколько цифр номера и обязательно заполнить поле "Юнит"!\n\n' \
           '- Как вывести на печать найденный отчёт?\n' \
           'На печать (на лист Excel) будет выведен отчёт, который развернули (кликнули по его названию). Для печати ' \
           'нескольких отчётов надо развернуть несколько интересующих отчётов. Каждый отчёт на отдельном листе ' \
           'Excel. Если не один отчёт не развёрнут, то на печать выведутся все найденные отчёты!\n\n' \
           '!!!Для администраторов!!!\n\n' \
           '- Как добавить данные?\n' \
           'Нажать на кнопку "Добавить" и выбрать один или несколько отчётов. При этом все поля и кнопки программы ' \
           'заблокируются, пока будет идти процесс загрузки новых данных.\n\n' \
           'После загрузки просмотреть файл "Log File" в соответствующей папке на наличие ошибок из-за которых ' \
           'данные не были загружены, а также просмотреть сводные данные.\n\n' \
           '- Что такое сводные данные?\n' \
           'Это сведения о номерах отчётов, номерах таблиц с данными в этих отчётах, их количестве и номерах юнитов. ' \
           'Достоверность этих и других данных ВАЖНЫ для корректной работы всей программы!!!\n\n' \
           '- Как удалять данные?\n' \
           'Удалять можно только таблицы целиком. Для этого поставьте флажок напротив интересующей таблицы и нажмите ' \
           'кнопку "Удалить".'
