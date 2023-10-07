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
                   'Diameter': ['ETER', 'INCH', 'ДИАМЕ', 'Día'.upper(), 'Ø'],
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
                   'Date': ['DATE'],
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
                    'Line': ['LINE', 'TAG', 'CONTR', 'OBJ', 'EQUIP'],
                    'Drawing': ['DRAW', 'ISOM'],
                    'Scanned_area': ['SCAN', 'AREA'],
                    'Item_description': ['ITEM', 'DESCR'],
                    'Diameter': ['ETER', 'INCH', 'ДИАМЕ', 'Ø', 'SIZE', 'DIMENS'],
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

# FAQ на русском языке
faq_text_ru = '- Для чего эта программа?\n' \
           'Для быстрого поиска измеренных значений и чертежей UTT и PAUT начиная с 2019 года.\n\n' \
           '- Как искать данные?\n' \
           'Заполнить одно или несколько полей для поиска. Для самого быстрого поиска рекомендуется всегда заполнять ' \
           'поле "Юнит"! Нет необходимости указывать полное название номеров линий, чертежей и т.д. достаточно указать последние ' \
           'несколько цифр номера и обязательно заполнить поле "Юнит"!\n' \
           'Например: для чертежей в формате ХХХХ-ХХХ (2920-002), для юнитов в формате ХХХ (320, 210), для номера репорта в ' \
           'формате ХХХ (060).\n\n' \
           '- Как вывести на печать найденный отчёт?\n' \
           'На печать (на лист Excel) будет выведен отчёт, который развернули (кликнули по его названию). Для печати ' \
           'нескольких отчётов надо развернуть несколько интересующих отчётов. Каждый отчёт будет выведен на отдельный лист ' \
           'Excel. По умолчанию отчёт сохраняется в папке "Reports for print".\n' \
           'Внимание!!! При печати репортов с длинным номером возможны предупреждения Excel - их можно игнорировать.\n\n' \
           '!!!Для администраторов!!!\n\n' \
           '- Как добавить данные?\n' \
           'Нажать на кнопку "Добавить" и выбрать один или несколько отчётов. При этом все поля и кнопки программы будут заблокированы, ' \
           'пока будет идти процесс загрузки новых данных.\n\n' \
           'После загрузки просмотреть файл "Log File" в соответствующей папке на наличие ошибок из-за которых данные не были загружены, ' \
           'а также просмотреть сводные данные.\n\n' \
           '- Что такое сводные данные?\n' \
           'Это сведения о номерах отчётов, номерах таблиц с данными в этих отчётах, их количестве и номерах юнитов.\n' \
           'По умолчанию сформированный отчёт сохраняется в папке "Statistic for print".\n'\
           'Достоверность этих и других данных ВАЖНЫ для корректной работы всей программы!!!\n\n' \
           '- Как удалять данные?\n' \
           'Удалять можно только таблицы целиком. Для этого поставьте флажок напротив интересующей таблицы и нажмите кнопку "Удалить".'

# FAQ на английском языке
faq_text_en = '- What is this program for??\n' \
           'For quick search of measured values and drawings UTT and PAUT from 2019.\n\n' \
           '- How to search for data?\n' \
           'Fill in one or more search fields. For the fastest search, it is recommended to always fill in ' \
           'field "Unit"! There is no need to specify the full name of line numbers, drawings, etc. enough to indicate the last ' \
           'a few digits of the number and be sure to fill in the "Unit" field!\n' \
           'For example: for drawings in the format XXX-XXX (2920-002), for units in the format XXX (320, 210), for the report number in ' \
           'format XXX (060).\n\n' \
           '- How to print the found report?\n' \
           'A report will be printed (on an Excel sheet) that has been expanded (clicked on its name). For print ' \
           'several reports, you need to deploy several reports of interest. Each report will be displayed on a separate sheet ' \
           'Excel. By default, the report is saved in the "Reports for print" folder.\n' \
           'Attention!!! When printing reports with a long number, Excel warnings are possible - they can be ignored.\n\n' \
           '!!!For administrators!!!\n\n' \
           '- How to add data?\n' \
           'Click on the "Add reports" button and select one or more reports. In this case, all fields and buttons of the program will be blocked, ' \
           'while the process of loading new data is in progress.\n\n' \
           'After downloading, view the "Log File" file in the appropriate folder for errors due to which the data was not downloaded, ' \
           'and view summary data.\n\n' \
           '- What is summary data?\n' \
           'This is information about report numbers, numbers of tables with data in these reports, their count and numbers of units.\n' \
           'By default, the generated report is saved in the "Statistic for print" folder.\n'\
           'The reliability of these and other data is IMPORTANT for the correct operation of the entire program!!!\n\n' \
           '- How to delete data?\n' \
           'You can only delete entire tables. To do this, check the box next to the table of interest and click the "Delete" button.'

# FAQ на казахском языке
faq_text_kz = '- Бұл бағдарлама не үшін?\n' \
           'Өлшенген мәндерді және UTT және PAUT сызбаларын жылдам алу үшін 2019 жылдан бастап\n\n' \
           '- Деректерді қалай іздеу керек?\n' \
           'Бір немесе бірнеше іздеу өрістерін толтырыңыз. Ең жылдам іздеу үшін әрқашан толтыру ұсынылады ' \
           '«Бірлік» өрісі! Жол нөмірлерінің, сызбалардың және т.б. толық атауын көрсетудің қажеті жоқ. соңғысын көрсету үшін жеткілікті ' \
           'Жол нөмірлерінің, сызбалардың және т.б. толык атауын корсетудін кажеті жок. soңgysyn korsetu ushin yetkіliktіn санның бірнеше цифрын ' \
           'және «Бірлік» өрісін міндетті түрде толтырыңыз.!\n' \
           'Мысалы: XXXX-XXX пішіміндегі сызбалар үшін (2920-002), XXXX форматындағы бірліктер үшін (320, 210), есеп нөмірі үшін ' \
           'XXX пішімі (060).\n\n' \
           '- Табылған есепті қалай басып шығаруға болады?\n' \
           'Кеңейтілген есеп басып шығарылады (Excel парағында) (атын басыңыз). Басып шығару үшін бірнеше есептер болса, бірнеше ' \
           'қызықты есептерді кеңейту қажет. Әрбір есеп жеке парақта көрсетіледі Excel. Әдепкі бойынша есеп «Reports for print» қалтасында ' \
           'сақталады.\n' \
           'Назар аударыңыз!!! Ұзын саны бар есептерді басып шығару кезінде Excel ескертулері мүмкін - оларды елемеу мүмкін.\n\n' \
           '!!!Әкімшілер үшін!!!\n\n' \
           '- Деректерді қалай қосуға болады?\n' \
           '«Қосу» түймесін басып, бір немесе бірнеше есепті таңдаңыз. Бұл жағдайда бағдарламаның барлық өрістері мен түймелері блокталады, ' \
           'жаңа деректерді жүктеу процесі жүріп жатқанда.\n\n' \
           'Жүктеп алғаннан кейін деректер жүктелмеген қателер үшін тиісті қалтадағы «Log File» қараңыз, және жиынтық деректерді қараңыз.\n\n' \
           '- Жиынтық деректер дегеніміз не?\n' \
           'Бұл есеп нөмірлері, осы есептердегі деректері бар кестелердің нөмірлері, олардың саны және бірлік нөмірлері туралы ақпарат.\n' \
           'Әдепкі бойынша жасалған есеп қалтада сақталады "Statistic for print".\n'\
           'Осы және басқа деректердің дәлдігі бүкіл бағдарламаның дұрыс жұмыс істеуі үшін МАҢЫЗДЫ.!!!\n\n' \
           '- Деректерді қалай жоюға болады?\n' \
           'Сіз тек толық кестелерді жоя аласыз. Ол үшін қызықтыратын кестенің жанындағы құсбелгіні қойып, «Жою» түймесін басыңыз.'

index_x11 = [800, 900, 1000, 1100, 1200, 1300, 1400, 1500, 1600, 1700, 1800, 1900, 2000, 2100, 2200, 2300, 2400, 2500, 2600, 2700, 2800, 2900, 3000]

# количество репортов, которое должно бы ть в каждой БД
count_reports_in_ndt = {'reports_db_OF_PAUT_19.sqlite': 0, 'reports_db_OF_PAUT_20.sqlite': 0, 'reports_db_OF_PAUT_21.sqlite': 0, 'reports_db_OF_PAUT_22.sqlite': 0,
                        'reports_db_OF_PAUT_23.sqlite': 0, 'reports_db_OF_PAUT_24.sqlite': 0, 'reports_db_OF_PAUT_25.sqlite': 0, 'reports_db_OF_PAUT_26.sqlite': 0,
                        'reports_db_OF_UTT_19.sqlite': 0, 'reports_db_OF_UTT_20.sqlite': 0, 'reports_db_OF_UTT_21.sqlite': 0, 'reports_db_OF_UTT_22.sqlite': 0,
                        'reports_db_OF_UTT_23.sqlite': 0, 'reports_db_OF_UTT_24.sqlite': 0, 'reports_db_OF_UTT_25.sqlite': 0, 'reports_db_OF_UTT_26.sqlite': 0,
                        'reports_db_ON_PAUT_19.sqlite': 0, 'reports_db_ON_PAUT_20.sqlite': 0, 'reports_db_ON_PAUT_21.sqlite': 353, 'reports_db_ON_PAUT_22.sqlite': 978,
                        'reports_db_ON_PAUT_23.sqlite': 0, 'reports_db_ON_PAUT_24.sqlite': 0, 'reports_db_ON_PAUT_25.sqlite': 0, 'reports_db_ON_PAUT_26.sqlite': 0,
                        'reports_db_ON_UTT_19.sqlite': 0, 'reports_db_ON_UTT_20.sqlite': 0, 'reports_db_ON_UTT_21.sqlite': 793, 'reports_db_ON_UTT_22.sqlite': 810,
                        'reports_db_ON_UTT_23.sqlite': 0, 'reports_db_ON_UTT_24.sqlite': 0, 'reports_db_ON_UTT_25.sqlite': 0, 'reports_db_ON_UTT_26.sqlite': 0,
                        'reports_db_OS_PAUT_19.sqlite': 0, 'reports_db_OS_PAUT_20.sqlite': 0, 'reports_db_OS_PAUT_21.sqlite': 11, 'reports_db_OS_PAUT_22.sqlite': 11,
                        'reports_db_OS_PAUT_23.sqlite': 0, 'reports_db_OS_PAUT_24.sqlite': 0, 'reports_db_OS_PAUT_25.sqlite': 0, 'reports_db_OS_PAUT_26.sqlite': 0,
                        'reports_db_OS_UTT_19.sqlite': 0, 'reports_db_OS_UTT_20.sqlite': 0, 'reports_db_OS_UTT_21.sqlite': 52, 'reports_db_OS_UTT_22.sqlite': 18,
                        'reports_db_OS_UTT_23.sqlite': 0, 'reports_db_OS_UTT_24.sqlite': 0, 'reports_db_OS_UTT_25.sqlite': 0, 'reports_db_OS_UTT_26.sqlite': 0}
