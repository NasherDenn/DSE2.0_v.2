# -*- coding: utf-8 -*-
# import time
# import hashlib

from PyQt5.QtSql import QSqlDatabase
from PyQt5.QtSql import QSqlQueryModel
from PyQt5.QtSql import QSqlTableModel
# from PyQt5 import QtWidgets

import YKR.utilities_interface
from YKR.utilities_interface import *
from YKR.add_reports import *
from YKR.props import *

import sys
from PyQt5.QtCore import Qt

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})

# настраиваем систему логирования
# дата (месяц, год) файла LogFile из системы
date_log_file = datetime.datetime.now().strftime("%m %Y")
# путь к папке где будет сохраняться LogFile
new_path_log_file = f'{os.path.abspath(os.getcwd())}\\Log File\\'
# если папка Log File не создана,
if not os.path.exists(new_path_log_file):
    # то создаём эту папку
    os.makedirs(new_path_log_file)
# путь сохранения LogFile
name_log_file = f'{new_path_log_file}{date_log_file} Log File.txt'
logging.basicConfig(level=logging.INFO,
                    handlers=[logging.FileHandler(filename=name_log_file, mode='a', encoding='utf-8')],
                    format='%(asctime)s [%(levelname)s] Пользователь: %(user)s - %(message)s', )
# дополняем базовый формат записи лог сообщения данными о пользователе
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})

logger_with_user.info('Запуск программы')

# создаём приложение
app = QApplication(sys.argv)
# создаём окно приложения
window = QWidget()
# название приложения
window.setWindowTitle('Data Search Engine')
# задаём стиль приложения Fusion
app.setStyle('Fusion')
# размер окна приложения
window.setFixedSize(1722, 965)

# устанавливаем favicon в окне приложения
icon = QIcon()
icon.addFile(f'{os.path.abspath(os.getcwd())}\\Images\\icon.ico', QSize(), QIcon.Normal, QIcon.Off)
icon.addFile(f'{os.path.abspath(os.getcwd())}\\Images\\icon.ico', QSize(), QIcon.Active, QIcon.On)

app.setWindowIcon(icon)

# задаём параметры стиля и оформления окна ввода
font = QFont()
font.setFamily(u"Arial")
font.setPointSize(14)
font.setItalic(False)

# создаём однострочное поле для ввода номера линии
line_search_line = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_line.setGeometry(QRect(181, 20, 561, 31))
# вывод надписи при наведении курсора
line_search_line.setToolTip('Полный или частичный номер линии/ёмкости, таговый номер')
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_line.setObjectName(u"line_search_line")
# дополнительные параметры
line_search_line.setFont(font)
line_search_line.setMouseTracking(False)
line_search_line.setContextMenuPolicy(Qt.NoContextMenu)
line_search_line.setAcceptDrops(True)
# line_search_line.setStyleSheet(u"")
line_search_line.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_line.setEchoMode(QLineEdit.Normal)
line_search_line.setCursorPosition(0)
line_search_line.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_line.setMaxLength(40)
line_search_line.setClearButtonEnabled(True)
# включаем переход фокуса по кнопке Tab или по клику мыши
line_search_line.setFocusPolicy(Qt.StrongFocus)
line_search_line.setText('')
line_search_line.setFocus()

# создаём однострочное поле для ввода номера чертежа
line_search_drawing = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_drawing.setGeometry(QRect(181, 60, 561, 31))
# вывод надписи при наведении курсора
line_search_drawing.setToolTip('Полный или частичный номер чертежа')
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_drawing.setObjectName(u"line_search_drawing")
# дополнительные параметры
line_search_drawing.setFont(font)
line_search_drawing.setMouseTracking(False)
line_search_drawing.setContextMenuPolicy(Qt.NoContextMenu)
line_search_drawing.setAcceptDrops(True)
# line_search_drawing.setStyleSheet(u"")
line_search_drawing.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_drawing.setEchoMode(QLineEdit.Normal)
line_search_drawing.setCursorPosition(0)
line_search_drawing.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_drawing.setMaxLength(40)
line_search_drawing.setClearButtonEnabled(True)
# включаем переход фокуса по кнопке Tab или по клику мыши
line_search_drawing.setFocusPolicy(Qt.StrongFocus)
line_search_drawing.setText('')


# создаём однострочное поле для ввода номера unit
line_search_unit = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_unit.setGeometry(QRect(181, 100, 170, 31))
# вывод надписи при наведении курсора
line_search_unit.setToolTip('Трёхзначный номер юнита')
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_unit.setObjectName(u"line_search_unit")
# дополнительные параметры
line_search_unit.setFont(font)
line_search_unit.setMouseTracking(False)
line_search_unit.setContextMenuPolicy(Qt.NoContextMenu)
line_search_unit.setAcceptDrops(True)
# line_search_unit.setStyleSheet(u"")
line_search_unit.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_unit.setEchoMode(QLineEdit.Normal)
line_search_unit.setCursorPosition(0)
line_search_unit.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_unit.setMaxLength(10)
line_search_unit.setClearButtonEnabled(True)
# включаем переход фокуса по кнопке Tab или по клику мыши
line_search_unit.setFocusPolicy(Qt.StrongFocus)
line_search_unit.setText('')

# создаём однострочное поле для ввода номера локации
line_search_item_description = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_item_description.setGeometry(QRect(521, 100, 220, 31))
# вывод надписи при наведении курсора
line_search_item_description.setToolTip('Полный(-ое) или частичный(-ое) номер/название локации')
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_item_description.setObjectName(u"line_search_item_description")
# дополнительные параметры
line_search_item_description.setFont(font)
line_search_item_description.setMouseTracking(False)
line_search_item_description.setContextMenuPolicy(Qt.NoContextMenu)
line_search_item_description.setAcceptDrops(True)
# line_search_item_description.setStyleSheet(u"")
line_search_item_description.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_item_description.setEchoMode(QLineEdit.Normal)
line_search_item_description.setCursorPosition(0)
line_search_item_description.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_item_description.setMaxLength(10)
line_search_item_description.setClearButtonEnabled(True)
# включаем переход фокуса по кнопке Tab или по клику мыши
line_search_item_description.setFocusPolicy(Qt.StrongFocus)
line_search_item_description.setText('')

# создаём однострочное поле для ввода номера репорта
line_search_number_report = QLineEdit(window)
# устанавливаем положение окна ввода и его размеры в родительском окне
line_search_number_report.setGeometry(QRect(181, 140, 561, 31))
# вывод надписи при наведении курсора
line_search_number_report.setToolTip('Полный или частичный номер отчёта')
# присваиваем уникальное объектное имя однострочному полю для ввода
line_search_number_report.setObjectName(u"line_search_number_report")
# дополнительные параметры
line_search_number_report.setFont(font)
line_search_number_report.setMouseTracking(False)
line_search_number_report.setContextMenuPolicy(Qt.NoContextMenu)
line_search_number_report.setAcceptDrops(True)
# line_search_number_report.setStyleSheet(u"")
line_search_number_report.setLocale(QLocale(QLocale.English, QLocale.UnitedStates))
line_search_number_report.setEchoMode(QLineEdit.Normal)
line_search_number_report.setCursorPosition(0)
line_search_number_report.setCursorMoveStyle(Qt.LogicalMoveStyle)
line_search_number_report.setMaxLength(40)
line_search_number_report.setClearButtonEnabled(True)
# включаем переход фокуса по кнопке Tab или по клику мыши
line_search_number_report.setFocusPolicy(Qt.StrongFocus)
line_search_number_report.setText('ut-22-010')

# создаём кнопку "Поиск"
button_search = QPushButton('Поиск', window)
# устанавливаем положение и размер кнопки для поиска в родительском окне (window)
button_search.setGeometry(760, 20, 161, 41)
# вывод надписи при наведении курсора
button_search.setToolTip('Поиск')
# присваиваем уникальное объектное имя кнопке "Поиск"
button_search.setObjectName(u"pushButton_search")
# задаём параметры стиля и оформления кнопки "Поиск"
font_button_search = QFont()
font_button_search.setFamily(u"Arial")
font_button_search.setPointSize(14)
button_search.setFont(font_button_search)
# включаем переход фокуса по кнопке Tab или по клику мыши
button_search.setFocusPolicy(Qt.StrongFocus)

# создаём кнопку печати
button_print = QPushButton('Печать', window)
# устанавливаем положение и размер кнопки печати в родительском окне (window)
button_print.setGeometry(QRect(760, 75, 161, 41))
# вывод надписи при наведении курсора
button_print.setToolTip('Печать')
# присваиваем уникальное объектное имя кнопке "Печать"
button_print.setObjectName(u"pushButton_print")
# задаём параметры стиля и оформления кнопки печати
font_button_print = QFont()
font_button_print.setFamily(u"Arial")
font_button_print.setPointSize(14)
button_print.setFont(font_button_print)
# дополнительные параметры
button_print.setFocusPolicy(Qt.ClickFocus)
# включаем переход фокуса по кнопке Tab
button_print.setFocusPolicy(Qt.TabFocus)

# создаём кнопку "Добавить"
button_add = QPushButton('Добавить', window)
# устанавливаем положение и размер кнопки "Добавить" в родительском окне (window)
button_add.setGeometry(QRect(760, 130, 161, 41))
# вывод надписи при наведении курсора
button_add.setToolTip('Добавить таблицы в БД')
# присваиваем уникальное объектное имя кнопке "Добавить"
button_add.setObjectName(u"pushButton_add")
# задаём параметры стиля и оформления кнопки "Добавить"
font_button_add = QFont()
font_button_add.setFamily(u"Arial")
font_button_add.setPointSize(14)
button_add.setFont(font_button_add)
# дополнительные параметры
button_add.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Добавить" до авторизации
button_add.setDisabled(True)
# включаем переход фокуса по кнопке Tab
button_add.setFocusPolicy(Qt.TabFocus)

# создаём кнопку "Закрыть" из программы
button_exit = QPushButton('Закрыть', window)
# устанавливаем положение и размер кнопки "Закрыть" для выхода из программы в родительском окне (window)
button_exit.setGeometry(QRect(1540, 904, 161, 41))
# вывод надписи при наведении курсора
button_exit.setToolTip('Выход из программы')
# присваиваем уникальное объектное имя кнопке "Закрыть"
button_exit.setObjectName(u"pushButton_exit")
# задаём параметры стиля и оформления кнопки "Закрыть"
font_button_exit = QFont()
font_button_exit.setFamily(u"Arial")
font_button_exit.setPointSize(14)
button_exit.setFont(font_button_exit)
# дополнительные параметры
button_exit.setFocusPolicy(Qt.ClickFocus)
# Закрытие программы при нажатии на кнопку "Закрыть"
button_exit.clicked.connect(qApp.exit)

# создаём кнопку "RU"
button_ru = QPushButton('RU', window)
# устанавливаем положение и размер кнопки "RU"
button_ru.setGeometry(QRect(1623, 20, 26, 26))
# вывод надписи при наведении курсора
button_ru.setToolTip('Русский язык')
button_ru.setCheckable(True)
button_ru.setChecked(True)

# создаём кнопку "EN"
button_en = QPushButton('EN', window)
# устанавливаем положение и размер кнопки "EN"
button_en.setGeometry(QRect(1649, 20, 26, 26))
# вывод надписи при наведении курсора
button_en.setToolTip('Английский язык')
# включаем возможность быть нажатой
button_en.setCheckable(True)
# по умолчанию не нажата
button_en.setChecked(False)

# создаём кнопку "KZ"
button_kz = QPushButton('KZ', window)
# устанавливаем положение и размер кнопки "KZ"
button_kz.setGeometry(QRect(1675, 20, 26, 26))
# вывод надписи при наведении курсора
button_kz.setToolTip('Казахский язык')
# включаем возможность быть нажатой
button_kz.setCheckable(True)
# по умолчанию не нажата
button_kz.setChecked(False)

# создаём кнопку "FAQ"
button_faq = QPushButton('FAQ', window)
# устанавливаем положение и размер кнопки "FAQ"
button_faq.setGeometry(QRect(1651, 186, 50, 50))
# вывод надписи при наведении курсора
button_faq.setToolTip('Часто задаваемые вопросы')
# по умолчанию не нажата
button_faq.setChecked(False)
# задаём параметры стиля и оформления кнопки "FAQ"
font_button_faq = QFont()
font_button_faq.setFamily(u"Arial")
font_button_faq.setPointSize(12)
button_faq.setFont(font_button_faq)


# создаём однострочное поле для ввода логина
line_login = QLineEdit(window)
# присваиваем уникальное объектное имя полю для ввода логина
line_login.setObjectName(u"line_login")
# устанавливаем положение и размер поля для ввода логина в родительском окне (window)
line_login.setGeometry(QRect(1470, 55, 111, 31))
# вывод надписи при наведении курсора
line_login.setToolTip('Поле для ввода логина')
# задаём параметры стиля и оформления поля для ввода логина
font_line_login = QFont()
font_line_login.setFamily(u"Arial")
font_line_login.setPointSize(11)
font_line_login.setItalic(True)
line_login.setFont(font_line_login)
# дополнительные параметры
line_login.setEchoMode(QLineEdit.Normal)
line_login.setMaxLength(10)
# устанавливаем исчезающий текст
line_login.setPlaceholderText('login')
line_login.setText('admin')

# создаём однострочное поле для ввода пароля
line_password = QLineEdit(window)
# присваиваем уникальное объектное имя полю для ввода пароля
line_password.setObjectName(u"line_password")
# устанавливаем положение и размер поля для ввода пароля в родительском окне (window)
line_password.setGeometry(QRect(1470, 95, 111, 31))
# вывод надписи при наведении курсора
line_password.setToolTip('Поле для ввода пароля')
# задаём параметры стиля и оформления поля для ввода пароля
font_line_password = QFont()
font_line_password.setFamily(u"Arial")
font_line_password.setPointSize(11)
font_line_password.setItalic(True)
line_password.setFont(font_line_password)
# дополнительные параметры
line_password.setEchoMode(QLineEdit.Password)
line_password.setMaxLength(10)
# устанавливаем исчезающий текст
line_password.setPlaceholderText('password')
line_password.setText('0751')

# устанавливаем надпись "Логин"
label_login = QLabel('Логин', window)
# присваиваем уникальное объектное имя надписи "Логин"
label_login.setObjectName(u"label_login")
# устанавливаем положение и размер поля для надписи "Логин" в родительском окне (window)
label_login.setGeometry(QRect(1400, 60, 61, 26))
# задаём параметры стиля и оформления поля для надписи "Логин"
font_label_login = QFont()
font_label_login.setFamily(u"Arial")
font_label_login.setPointSize(12)
font_label_login.setItalic(True)
label_login.setFont(font_label_login)

# устанавливаем надпись "Пароль"
label_password = QLabel('Пароль', window)
# присваиваем уникальное объектное имя надписи "Пароль"
label_password.setObjectName(u"label_password")
# устанавливаем положение и размер поля для надписи "Пароль" в родительском окне (window)
label_password.setGeometry(QRect(1390, 100, 81, 26))
# задаём параметры стиля и оформления поля для надписи "Пароль"
font_label_password = QFont()
font_label_password.setFamily(u"Arial")
font_label_password.setPointSize(12)
font_label_password.setItalic(True)
# скрываем введённые с клавиатуры символы при вводе в поле "Пароль"
label_password.setFont(font_label_password)

# устанавливаем надпись "Линия"
label_line = QLabel('Линия/Ёмкость', window)
# присваиваем уникальное объектное имя надписи "Линия"
label_line.setObjectName(u"label_line")
# устанавливаем положение и размер поля для надписи "Линия" в родительском окне (window)
label_line.setGeometry(QRect(15, 25, 156, 26))
# задаём параметры стиля и оформления поля для надписи "Линия"
font_label_line = QFont()
font_label_line.setFamily(u"Arial")
font_label_line.setPointSize(12)
font_label_line.setItalic(True)
label_line.setFont(font_label_line)
label_line.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Чертёж"
label_drawing = QLabel('Номер чертежа', window)
# присваиваем уникальное объектное имя надписи "Чертёж"
label_drawing.setObjectName(u"label_drawing")
# устанавливаем положение и размер поля для надписи "Чертёж" в родительском окне (window)
label_drawing.setGeometry(QRect(20, 65, 151, 26))
# задаём параметры стиля и оформления поля для надписи "Чертёж"
font_label_drawing = QFont()
font_label_drawing.setFamily(u"Arial")
font_label_drawing.setPointSize(12)
font_label_drawing.setItalic(True)
label_drawing.setFont(font_label_drawing)
label_drawing.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Юнит"
label_unit = QLabel('Юнит', window)
# присваиваем уникальное объектное имя надписи "Юнит"
label_unit.setObjectName(u"label_label_unit")
# устанавливаем положение и размер поля для надписи "Юнит" в родительском окне (window)
label_unit.setGeometry(QRect(20, 105, 151, 26))
# задаём параметры стиля и оформления поля для надписи "Юнит"
font_label_unit = QFont()
font_label_unit.setFamily(u"Arial")
font_label_unit.setPointSize(12)
font_label_unit.setItalic(True)
label_unit.setFont(font_label_unit)
label_unit.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Номер локации"
label_item_description = QLabel('Номер локации', window)
# присваиваем уникальное объектное имя надписи "Номер локации"
label_item_description.setObjectName(u"label_item_description")
# устанавливаем положение и размер поля для надписи "Номер локации" в родительском окне (window)
label_item_description.setGeometry(QRect(360, 105, 151, 26))
# задаём параметры стиля и оформления поля для надписи "Номер локации"
font_label_item_description = QFont()
font_label_item_description.setFamily(u"Arial")
font_label_item_description.setPointSize(12)
font_label_item_description.setItalic(True)
label_item_description.setFont(font_label_item_description)
label_item_description.setAlignment(Qt.AlignRight)

# устанавливаем надпись "Номер отчёта"
label_number_report = QLabel('Номер отчёта', window)
# присваиваем уникальное объектное имя надписи "Номер локации"
label_number_report.setObjectName(u"label_number_report")
# устанавливаем положение и размер поля для надписи "Номер локации" в родительском окне (window)
label_number_report.setGeometry(QRect(20, 145, 151, 26))
# задаём параметры стиля и оформления поля для надписи "Номер локации"
font_label_number_report = QFont()
font_label_number_report.setFamily(u"Arial")
font_label_number_report.setPointSize(12)
font_label_number_report.setItalic(True)
label_number_report.setFont(font_label_number_report)
label_number_report.setAlignment(Qt.AlignRight)

# создаём кнопку "Войти"
button_log_in = QPushButton('Войти', window)
# устанавливаем положение и размер кнопки "Войти" в родительском окне (window)
button_log_in.setGeometry(QRect(1590, 55, 111, 31))
# вывод надписи при наведении курсора
button_log_in.setToolTip('Войти в систему')
# присваиваем уникальное объектное имя кнопке "Войти"
button_log_in.setObjectName(u"pushButton_enter")
# задаём параметры стиля и оформления кнопки "Войти"
font_button_log_in = QFont()
font_button_log_in.setFamily(u"Arial")
font_button_log_in.setPointSize(14)
button_log_in.setFont(font_button_log_in)
# дополнительные параметры
button_log_in.setFocusPolicy(Qt.ClickFocus)

# создаём кнопку "Выйти"
button_log_out = QPushButton('Выйти', window)
# устанавливаем положение и размер кнопки "Выйти" в родительском окне (window)
button_log_out.setGeometry(QRect(1590, 95, 111, 31))
# вывод надписи при наведении курсора
button_log_out.setToolTip('Выйти из системы')
# присваиваем уникальное объектное имя кнопке "Выйти"
button_log_out.setObjectName(u"pushButton_out")
# задаём параметры стиля и оформления кнопки "Выйти"
font_button_log_out = QFont()
font_button_log_out.setFamily(u"Arial")
font_button_log_out.setPointSize(14)
button_log_out.setFont(font_button_log_out)
# дополнительные параметры
button_log_out.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Выйти" до авторизации
button_log_out.setDisabled(True)

# создаём кнопку "Удалить таблицу"
button_delete_table = QPushButton('Удалить таблицу', window)
# устанавливаем положение и размер кнопки "Удалить таблицу" в родительском окне (window)
button_delete_table.setGeometry(QRect(20, 904, 191, 41))
# вывод надписи при наведении курсора
button_delete_table.setToolTip('Удалить выбранную таблицу целиком')
# присваиваем уникальное объектное имя кнопке "Удалить таблицу"
button_delete_table.setObjectName(u"pushButton_delete")
# задаём параметры стиля и оформления кнопки "Удалить таблицу"
font_button_delete_table = QFont()
font_button_delete_table.setFamily(u"Arial")
font_button_delete_table.setPointSize(14)
button_delete_table.setFont(font_button_delete_table)
# дополнительные параметры
button_delete_table.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Удалить" до авторизации
button_delete_table.setDisabled(True)

# создаём кнопку "Удалить строку"
button_delete_row = QPushButton('Удалить строку', window)
# устанавливаем положение и размер кнопки "Удалить строку" в родительском окне (window)
button_delete_row.setGeometry(QRect(231, 904, 171, 41))
# вывод надписи при наведении курсора
button_delete_row.setToolTip('Удалить строку из таблицы, на которой находится курсор')
# присваиваем уникальное объектное имя кнопке "Удалить строку"
button_delete_row.setObjectName(u"pushButton_delete_row")
# задаём параметры стиля и оформления кнопки "Удалить строку"
font_button_delete_row = QFont()
font_button_delete_row.setFamily(u"Arial")
font_button_delete_row.setPointSize(14)
button_delete_row.setFont(font_button_delete_row)
# дополнительные параметры
button_delete_row.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Удалить строку" до авторизации
button_delete_row.setDisabled(True)

# создаём кнопку "Добавить строку"
button_add_row = QPushButton('Добавить строку', window)
# устанавливаем положение и размер кнопки "Добавить строку" в родительском окне (window)
button_add_row.setGeometry(QRect(495, 904, 191, 41))
# вывод надписи при наведении курсора
button_add_row.setToolTip('Добавить строку в конец таблицы')
# присваиваем уникальное объектное имя кнопке "Добавить строку"
button_add_row.setObjectName(u"pushButton_button_add_row")
# задаём параметры стиля и оформления кнопки "Добавить строку"
font_button_add_row = QFont()
font_button_add_row.setFamily(u"Arial")
font_button_add_row.setPointSize(14)
button_add_row.setFont(font_button_add_row)
# дополнительные параметры
button_add_row.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Добавить строку" до авторизации
button_add_row.setDisabled(True)

# создаём кнопку "Сохранить"
button_save = QPushButton('Сохранить', window)
# устанавливаем положение и размер кнопки "Сохранить" в родительском окне (window)
button_save.setGeometry(QRect(706, 904, 191, 41))
# вывод надписи при наведении курсора
button_save.setToolTip('Сохранить изменения')
# присваиваем уникальное объектное имя кнопке "Сохранить"
button_save.setObjectName(u"pushButton_button_save")
# задаём параметры стиля и оформления кнопки "Сохранить"
font_button_save = QFont()
font_button_save.setFamily(u"Arial")
font_button_save.setPointSize(14)
button_save.setFont(font_button_save)
# дополнительные параметры
button_save.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Сохранить" до авторизации
button_save.setDisabled(True)

# создаём кнопку "Верификация"
button_verification = QPushButton('Верификация', window)
# устанавливаем положение и размер кнопки "Верификация" в родительском окне (window)
button_verification.setGeometry(QRect(990, 904, 191, 41))
# вывод надписи при наведении курсора
button_verification.setToolTip('Автоматическая проверка данных в БД')
# присваиваем уникальное объектное имя кнопке "Верификация"
button_verification.setObjectName(u"pushButton_verification")
# задаём параметры стиля и оформления кнопки "Верификация"
font_button_verification = QFont()
font_button_verification.setFamily(u"Arial")
font_button_verification.setPointSize(14)
button_verification.setFont(font_button_verification)
# дополнительные параметры
button_verification.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Верификация" до авторизации
button_verification.setDisabled(True)

# создаём кнопку "Сводные данные"
button_statistic_master = QPushButton('Сводные данные', window)
# устанавливаем положение и размер кнопки "Сводные данные" в родительском окне (window)
button_statistic_master.setGeometry(QRect(1201, 904, 191, 41))
# вывод надписи при наведении курсора
button_statistic_master.setToolTip('Общие данные из БД')
# присваиваем уникальное объектное имя кнопке "Сводные данные"
button_statistic_master.setObjectName(u"pushButton_statistic_master")
# задаём параметры стиля и оформления кнопки "Сводные данные"
font_button_statistic_master = QFont()
font_button_statistic_master.setFamily(u"Arial")
font_button_statistic_master.setPointSize(14)
button_statistic_master.setFont(font_button_statistic_master)
# дополнительные параметры
button_statistic_master.setFocusPolicy(Qt.ClickFocus)
# делаем неактивной кнопку "Сводные данные" до авторизации
button_statistic_master.setDisabled(True)

# вставляем картинку YKR
label_ykr = QLabel(window)
label_ykr.setObjectName(u"YKR")
label_ykr.setGeometry(QRect(974, 20, 111, 121))
# label_ykr.setGeometry(QRect(990, 10, 111, 121))
# вывод надписи при наведении курсора
label_ykr.setToolTip('Yeskert Kyzmet Rutledge')
label_ykr.setPixmap(QPixmap(f'{os.path.abspath(os.getcwd())}\\Images\\logo_ykr.png'))

# вставляем картинку NCA
label_nca = QLabel(window)
label_nca.setObjectName(u"NCA")
# вывод надписи при наведении курсора
label_nca.setToolTip('National Accreditation Center')
label_nca.setGeometry(QRect(1105, 20, 111, 121))
# label_nca.setGeometry(QRect(1120, 10, 111, 121))
label_nca.setPixmap(QPixmap(f'{os.path.abspath(os.getcwd())}\\Images\\logo_nca.png'))

# вставляем картинку NCOC
label_ncoc = QLabel(window)
label_ncoc.setObjectName(u"NCOC")
# вывод надписи при наведении курсора
label_ncoc.setToolTip('North Caspian Operating Company')
label_ncoc.setGeometry(QRect(1236, 20, 111, 121))
# label_ncoc.setGeometry(QRect(1250, 13, 111, 115))
label_ncoc.setPixmap(QPixmap(f'{os.path.abspath(os.getcwd())}\\Images\\logo_ncoc.png'))

# общая область с боковой полосой прокрутки
scroll_area = QScrollArea(window)
scroll_area.setObjectName(u'Scroll_Area')
# полоса прокрутки появляется, только если таблицы больше самой области прокрутки
scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
# задаём размер области с полосой прокрутки
scroll_area.setGeometry(20, 245, 1680, 650)

# создаём группу из чек-бокса 'ON', 'OF', 'OS'
groupBox_location = QGroupBox(window)
groupBox_location.setObjectName(u"groupBox_radio")
# устанавливаем размер группы чек-бокса
groupBox_location.setGeometry(QRect(20, 180, 161, 56))
# устанавливаем название группы чек-бокса
groupBox_location.setTitle('Локация')
groupBox_location.setStyleSheet('''QGroupBox {border: 0.5px solid grey;};
                                   QGroupBox:title{
                                   subcontrol-origin: margin;
                                   # subcontrol-origin: margin;
                                   subcontrol-position: top center;
                                   padding: 0 3px 0 3px;
                                }''')

# создаём чек-бокс локации 'ON'
checkBox_on = QCheckBox(groupBox_location)
checkBox_on.setObjectName(u"checkBox_on")
# устанавливаем положение внутри группы
checkBox_on.setGeometry(QRect(10, 25, 45, 20))
# вывод надписи при наведении курсора
checkBox_on.setToolTip('Onshore')
# указываем текст чек-бокса
checkBox_on.setText('ON')
checkBox_on.setChecked(True)

# создаём чек-бокс локации 'OF'
checkBox_of = QCheckBox(groupBox_location)
checkBox_of.setObjectName(u"checkBox_of")
# устанавливаем положение внутри группы
checkBox_of.setGeometry(QRect(61, 25, 42, 20))
# вывод надписи при наведении курсора
checkBox_of.setToolTip('Offshore')
# указываем текст чек-бокса
checkBox_of.setText('OF')
# временно, пока не загрузятся другие локации
checkBox_of.setEnabled(True)

# создаём чек-бокс локации 'OS'
checkBox_os = QCheckBox(groupBox_location)
checkBox_os.setObjectName(u"checkBox_os")
# устанавливаем положение внутри группы
checkBox_os.setGeometry(QRect(110, 25, 42, 20))
# вывод надписи при наведении курсора
checkBox_os.setToolTip('Off Side Shore')
# указываем текст чек-бокса
checkBox_os.setText('OS')
# временно, пока не загрузятся другие локации
checkBox_os.setEnabled(True)

# создаём группу из чек-боксов методов контроля
groupBox_ndt = QGroupBox(window)
groupBox_ndt.setObjectName(u"groupBox_ndt")
groupBox_ndt.setGeometry(QRect(190, 180, 161, 56))
# устанавливаем название группы чек-боксов
groupBox_ndt.setTitle('Метод контроля')
groupBox_ndt.setStyleSheet('''QGroupBox {border: 0.5px solid grey;};
                              QGroupBox:title{
                              subcontrol-origin: margin;
                              subcontrol-position: top center;
                              padding: 0 3px 0 3px;
                           }''')

# создаём чек-бокс метода контроля 'UTT'
checkBox_utt = QCheckBox(groupBox_ndt)
checkBox_utt.setObjectName(u"checkBox_utt")
# устанавливаем положение внутри группы
checkBox_utt.setGeometry(QRect(10, 25, 61, 20))
# вывод надписи при наведении курсора
checkBox_utt.setToolTip('Ultrasonic Testing Thickness')
# указываем текст чек-бокса
checkBox_utt.setText('UTT')
# делаем чек-бокс 'UTT' активным по умолчанию
checkBox_utt.setChecked(True)

# создаём чек-бокс метода контроля 'PAUT'
checkBox_paut = QCheckBox(groupBox_ndt)
checkBox_paut.setObjectName(u"checkBox_paut")
# устанавливаем положение внутри группы
checkBox_paut.setGeometry(QRect(80, 25, 61, 20))
# вывод надписи при наведении курсора
checkBox_paut.setToolTip('Phased Array Ultrasonic Testing')
# указываем текст чек-бокса
checkBox_paut.setText('PAUT')

# создаём группу из чек-боксов годов
groupBox_year = QGroupBox(window)
groupBox_year.setObjectName(u"groupBox_year")
# устанавливаем размер группы радио-кнопок
groupBox_year.setGeometry(QRect(360, 180, 561, 56))
# вывод надписи при наведении курсора
groupBox_year.setToolTip('Год контроля')
# устанавливаем название группы чек-боксов
groupBox_year.setTitle('Год контроля')
groupBox_year.setStyleSheet('''QGroupBox {border: 0.5px solid grey;};
                               QGroupBox:title{
                               subcontrol-origin: margin;
                               subcontrol-position: top center;
                               padding: 0 3px 0 3px;
                            }''')

# создаём чек-бокс года '2024'
checkBox_2024 = QCheckBox(groupBox_year)
checkBox_2024.setObjectName(u"checkBox_2024")
# устанавливаем положение внутри группы
checkBox_2024.setGeometry(QRect(10, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2024.setText('2024')
# временно, пока не загрузятся другие локации
checkBox_2024.setEnabled(True)

# создаём чек-бокс года '2023'
checkBox_2023 = QCheckBox(groupBox_year)
checkBox_2023.setObjectName(u"checkBox_2023")
# устанавливаем положение внутри группы
checkBox_2023.setGeometry(QRect(80, 25, 61, 20))
# checkBox_2023.setGeometry(QRect(10, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2023.setText('2023')
# временно, пока не загрузятся другие локации
checkBox_2023.setEnabled(True)

# создаём чек-бокс года '2022'
checkBox_2022 = QCheckBox(groupBox_year)
checkBox_2022.setObjectName(u"checkBox_2022")
# устанавливаем положение внутри группы
checkBox_2022.setGeometry(QRect(150, 25, 61, 20))
# checkBox_2022.setGeometry(QRect(80, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2022.setText('2022')
# делаем чек-бокс '2022' активным по умолчанию
checkBox_2022.setChecked(True)

# создаём чек-бокс года '2021'
checkBox_2021 = QCheckBox(groupBox_year)
checkBox_2021.setObjectName(u"checkBox_2021")
# устанавливаем положение внутри группы
checkBox_2021.setGeometry(QRect(220, 25, 61, 20))
# checkBox_2021.setGeometry(QRect(150, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2021.setText('2021')
# временно, пока не загрузятся другие локации
checkBox_2021.setEnabled(True)

# создаём чек-бокс года '2020'
checkBox_2020 = QCheckBox(groupBox_year)
checkBox_2020.setObjectName(u"checkBox_2020")
# устанавливаем положение внутри группы
checkBox_2020.setGeometry(QRect(290, 25, 61, 20))
# checkBox_2020.setGeometry(QRect(220, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2020.setText('2020')
# временно, пока не загрузятся другие локации
checkBox_2020.setEnabled(True)

# создаём чек-бокс года '2019'
checkBox_2019 = QCheckBox(groupBox_year)
checkBox_2019.setObjectName(u"checkBox_2019")
# устанавливаем положение внутри группы
checkBox_2019.setGeometry(QRect(360, 25, 61, 20))
# checkBox_2019.setGeometry(QRect(290, 25, 61, 20))
# указываем текст чек-бокса
checkBox_2019.setText('2019')
# временно, пока не загрузятся другие локации
checkBox_2019.setEnabled(True)

# метка авторизации
authorization = False


# нажатие кнопки "Войти"
def log_in():
    # если ничего не введено в поля "Логин" и "Пароль"
    if line_login.text() == '' and line_password.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы ничего не ввели в поля "Логин" и "Пароль"!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error('Попытка авторизоваться - не заполнены поля "Логин" и "Пароль"')
    # если ничего не введено в поле "Логин"
    elif line_login.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не заполнили поле "Логин"!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error(
            'Попытка авторизоваться - не заполнено поле "Логин", указан пароль - "{}"'.format(line_password.text()))
    # если ничего не введено в поле "Пароль"
    elif line_password.text() == '':
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не заполнили поле "Пароль"!!!',
            buttons=QMessageBox.Ok
        )
        logger_with_user.error(
            'Попытка авторизоваться - Не заполнено поле "Пароль", указан логин - "{}"'.format(line_login.text()))
    # если правильно введён логин и пароль
    elif line_login.text() == 'admin' and line_password.text() == '0751':
        # делаем активными кнопки "Добавить", "Удалить", "Выйти", "Сводные данные", "Верификация"
        button_delete_table.setDisabled(False)
        button_delete_row.setDisabled(False)
        button_add_row.setDisabled(False)
        button_save.setDisabled(False)
        button_log_out.setDisabled(False)
        button_add.setDisabled(False)
        button_statistic_master.setDisabled(False)
        button_verification.setDisabled(False)
        # очищаем поле ввода логина и пароля
        line_login.clear()
        line_password.clear()
        # блокируем кнопку "Войти"
        button_log_in.setDisabled(True)
        logger_with_user.info('Пользователь авторизовался')
        # сигнал о том, что выполнена авторизация
        global authorization
        authorization = True

        # если отображаются найденные таблицы
        if window.findChildren(QTableView):
            open_scroll_area = window.findChildren(QScrollArea)
            for scroll in open_scroll_area:
                open_push_button = scroll.findChildren(QPushButton)
                # список номеров кнопок, которые являются кнопками чертежей
                button_drawing = []
                # сдвигаем кнопки названий найденных таблиц и чертежей в бок для отображения флажков при авторизации
                for list_push in all_list_button_for_drawing:
                    x11 = 920
                    for push in list_push:
                        for index_push_button, push_button in enumerate(open_push_button):
                            if push_button == push:
                                push_button.move(x11, push_button.y())
                                push_button.repaint()
                                x11 += 85
                                button_drawing.append(index_push_button)
                for i in range(len(open_push_button)):
                    if i not in button_drawing:
                        open_push_button[i].move(20, open_push_button[i].y())
                        open_push_button[i].repaint()

                open_check_box = scroll.findChildren(QCheckBox)
                # делаем видимые флажки
                for check_box in open_check_box:
                    check_box.show()
    # если неправильно введён логин или пароль
    else:
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы ввели не правильный логин или пароль!!!',
            buttons=QMessageBox.Ok)
        logger_with_user.error(
            'Попытка авторизоваться - Введён неверный логин "{}" или пароль "{}"'.format(line_login.text(), line_password.text()))


# нажатие кнопки "Выйти"
def log_out():
    # делаем НЕ активными кнопки "Добавить", "Удалить", "Выйти", "Сводные данные", "Выйти", "Верификация"
    button_delete_table.setDisabled(True)
    button_delete_row.setDisabled(True)
    button_add_row.setDisabled(True)
    button_save.setDisabled(True)
    button_log_out.setDisabled(True)
    button_add.setDisabled(True)
    button_statistic_master.setDisabled(True)
    button_verification.setDisabled(True)
    # разблокируем кнопку "Войти"
    button_log_in.setDisabled(False)
    logger_with_user.info('Пользователь вышел')
    # сбрасываем на ноль авторизацию для отображения флажков
    global authorization
    authorization = False

    # если отображаются найденные таблицы
    if window.findChildren(QTableView):
        open_scroll_area = window.findChildren(QScrollArea)

        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            # список номеров кнопок, которые являются кнопками чертежей
            button_drawing = []
            # сдвигаем кнопки названий найденных таблиц и чертежей в бок для отображения флажков при авторизации
            for list_push in all_list_button_for_drawing:
                x11 = 900
                for push in list_push:
                    for index_push_button, push_button in enumerate(open_push_button):
                        if push_button == push:
                            push_button.move(x11, push_button.y())
                            push_button.repaint()
                            x11 += 85
                            button_drawing.append(index_push_button)
            for i in range(len(open_push_button)):
                if i not in button_drawing:
                    open_push_button[i].move(0, open_push_button[i].y())
                    open_push_button[i].repaint()

            open_check_box = scroll.findChildren(QCheckBox)
            # скрываем флажки
            for check_box in open_check_box:
                check_box.close()


# словарь выбранных фильтров (локация, метод, год) для поиска
data_filter_for_search = dict()
data_filter_for_search['location'] = {'ON': True, 'OF': False, 'OS': False}
data_filter_for_search['method'] = {'UTT': True, 'PAUT': False}
data_filter_for_search['year'] = {'2019': False, '2020': False, '2021': False, '2022': True, '2023': False}


# обработчик события выбора одной из локаций (on, of, os)
def on_button_clicked_location():
    check_button_location = QObject().sender()
    data_filter_for_search['location'][check_button_location.text()] = check_button_location.isChecked()


# определяем какая локация выбрана
for button in groupBox_location.findChildren(QCheckBox):
    button.clicked.connect(on_button_clicked_location)


# обработчик события выбора одного из методов контроля (utt, paut)
def on_button_clicked_ndt():
    check_button_ndt = QObject().sender()
    data_filter_for_search['method'][check_button_ndt.text()] = check_button_ndt.isChecked()


# определяем какой(-ие) методы контроля выбраны
for button in groupBox_ndt.findChildren(QCheckBox):
    button.clicked.connect(on_button_clicked_ndt)


# обработчик события выбора одного из фильтров (локация, метод, год)
def on_button_clicked_year():
    check_button_year = QObject().sender()
    data_filter_for_search['year'][check_button_year.text()] = check_button_year.isChecked()


# определяем какой(-ие) года выбраны
for button in groupBox_year.findChildren(QCheckBox):
    button.clicked.connect(on_button_clicked_year)

# список найденных данных
list_sqm = []
# список областей для вывода таблиц
list_table_view = []
# список всех наёденных чертежей в рамках одного запроса
all_list_button_for_drawing = []
# активатор остановки изменения ширины фрейма, если хоть в одном найденном репорте больше, чем девять чертежй
# stop_width_frame = False


# нажатие на кнопку "Поиск"
def search():
    # список найденных данных
    global list_sqm, all_list_button_for_drawing, list_table_view
    list_sqm = []
    list_table_view = []
    all_list_button_for_drawing = []

    # сбрасывание ширины фрейма на 1680
    YKR.utilities_interface.new_width_frame = 1680

    # проверяем наличие областей tableView для вывода данных и кнопок в ней
    # если есть, то закрываем их, чтобы не наслаивались
    if window.findChildren(QTableView):
        open_tableview = window.findChildren(QTableView)
        for tableview in open_tableview:
            tableview.setParent(None)
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            for push_button in open_push_button:
                # то сдвигаем вбок кнопки названий таблиц
                push_button.setParent(None)
    # проверяем какой язык выбран
    if button_ru.isChecked():
        language = 'ru'
    if button_en.isChecked():
        language = 'en'
    if button_kz.isChecked():
        language = 'kz'
    # словарь введённых данных в поля для поиска
    data_for_search = dict()
    # определяем какие данные для поиска введены в поля для поиска
    data_for_search['line_search'] = [line_search_line.text()]
    data_for_search['drawing_search'] = [line_search_drawing.text()]
    data_for_search['unit'] = [line_search_unit.text()]
    data_for_search['item_description_search'] = [line_search_item_description.text()]
    data_for_search['number_report_search'] = [line_search_number_report.text()]
    # проверка - введены ли данные для поиска и выбраны ли все фильтры
    if not_check_data_and_filter(data_filter_for_search, data_for_search):
        QMessageBox.information(
            window,
            'Внимание!',
            'Вы не ввели данные для поиска или не выбрали не один фильтр',
            buttons=QMessageBox.Ok
        )
        return

    # замораживаем все кнопки и поля на время поиска данных
    freeze_button()
    search_picture = Searching()
    search_picture.start_loading()
    window.repaint()

    # собираем названия БД, по выбранным фильтрам, в которых надо искать данные
    db_for_search = define_db_for_search(data_filter_for_search)
    # получаем значения из полей для ввода
    values_for_search = dict()
    values_for_search['line'] = line_search_line.text()
    values_for_search['drawing'] = line_search_drawing.text()
    values_for_search['unit'] = line_search_unit.text()
    values_for_search['item_description'] = line_search_item_description.text()
    values_for_search['number_report'] = line_search_number_report.text()

    # ищем таблицы в БД для дальнейшего вывода
    find_data = look_up_data(db_for_search, values_for_search)
    # frame в который будут вставляться, таблицы чтобы при большом количестве таблиц появлялась полоса прокрутки
    frame_for_table = QFrame()
    # помещаем frame в область с полосой прокрутки
    scroll_area.setWidget(frame_for_table)
    # начальная точка отсчёта по Y для вывода данных
    y1 = -20
    # счётчик общего количества всех строк при поиске
    all_count_row_in_search = 0
    # счётчик общего количества всех таблиц с искомыми данными
    all_count_table_in_search = 0
    # # список областей для вывода таблиц
    # list_table_view = []
    # список всех найденных таблиц
    list_button_for_table = []
    # словарь последовательностей пути и чертежей
    # dict_path_draw = {}
    list_height_table_view = []
    list_check_box = []
    # find_data[0] - основные данные
    # сортируем года в порядке убывания
    sort_db_year = sort_year(find_data[0])
    # индекс для последовательности путей и чертежей
    # index_path_draw = 0
    for db in sort_db_year:
        # сортируем таблицы в году по датам в обратном порядке, что бы вверху отображались более ранние отчёты
        unsort = sort_date(find_data[0][db].keys())
        sort = sorted(unsort, reverse=True)
        for key_date in sort:
            # итерируемся по несортированному начальному словарю с данными, на основании отсортированных дат
            for table in find_data[0][db][unsort[key_date]]:
                # создаём соединение с базой данной
                con = QSqlDatabase.addDatabase('QSQLITE')
                # передаём имя базы данных для открытия
                con.setDatabaseName(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
                if not con.open():
                    QMessageBox.critical(
                        None,
                        'App name Error',
                        'Error to connect to the database')
                    logger_with_user.error('Отсутствует соединение с базой данных')
                    sys.exit()
                con.open()
                # задаём поле для вывода данных из базы данных, размещённую в области с полосой прокрутки
                table_index = QTableView(frame_for_table)

                # создаём модель
                if authorization:
                    sqm = QSqlTableModel(parent=window)
                    sqm.setEditStrategy(QSqlTableModel.OnManualSubmit)
                else:
                    sqm = QSqlQueryModel(parent=window)

                # устанавливаем ширину столбцов под содержимое
                table_index.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                # устанавливаем высоту столбцов под содержимое
                table_index.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
                # устанавливаем разный цвет фона для чётных и нечётных строк
                table_index.setAlternatingRowColors(True)
                # помещаем запрос в поле для вывода данных
                table_index.setModel(sqm)
                # find_data[1] - path (путь поиска - 1, 2, 3, ...)
                # find_data[2] - place_for_search (поля для поиска, которые заполнены)
                # find_data[3] - values (введённые значения в поля для поиска)
                # find_data[4] - list_date (список дат каждого репорта)

                if find_data[1] == 1:
                    # делаем запрос в модели
                    if authorization:
                        sqm.setTable(table)
                        sqm.select()

                    else:
                        sqm.setQuery('SELECT * FROM {}'.format(table))
                    # количество строк в найденной таблице
                    count_row_in_table = sqm.rowCount()
                    # сумма всех строк при поиске
                    all_count_row_in_search += count_row_in_table
                    # количество таблиц, в которых найдены данные
                    all_count_table_in_search += 1
                    # высота одной таблицы tableView = количество строк в одной таблице * высоту одной строки + высота строки названия столбцов
                    table_height_for_data_output = count_row_in_table * 25 + 25
                    table_index.hide()
                    y1 += 20
                    # список чертежей в рамках одного репорта
                    list_button_for_drawing = []
                    table_index.setGeometry(QRect(0, y1, 2500, table_height_for_data_output))

                    if sqm.rowCount() != 0:
                        list_sqm.append(sqm)
                        # номер линии, сосуда для кнопки названия таблицы
                        line_for_table_name_buttons = sqm.query().value('line')
                        button_for_table = table_name_buttons(frame_for_table, y1, authorization, table, unsort[key_date], language,
                                                              line_for_table_name_buttons)
                        # папка с чертежами для данного репорта
                        dir_with_drawing = table.replace('_', '-')[3:]
                        dir_drawings_db = db.replace('reports', 'drawings')[:-7]
                        # путь, где лежат чертежи для репорта
                        path_drawing = f'{os.path.abspath(os.getcwd())}\\Drawings\\{dir_drawings_db}\\{dir_with_drawing}'
                        # список чертежей для данного репорта
                        try:
                            list_drawing_button = os.listdir(path_drawing)
                        except FileNotFoundError:
                            list_drawing_button = []
                            logger_with_user.info(f'Отсутствуют чертежи для отчёта {dir_with_drawing} в базе чертежей {dir_drawings_db}')
                        # выставляем координату от левого края
                        # если авторизовались
                        if authorization:
                            x11 = 920
                        # если нет
                        else:
                            x11 = 900
                        # итерируемся по количеству чертежей для создания нужного количества кнопок номеров чертежей
                        if len(list_drawing_button) > 0:
                            for index, draw in enumerate(list_drawing_button):
                                # создаём кнопку для чертежа
                                button_for_drawing = drawing_name_buttons(frame_for_table, y1, x11, language, index, path_drawing, draw)
                                x11 += 110
                                list_button_for_drawing.append(button_for_drawing)
                        all_list_button_for_drawing.append(list_button_for_drawing)
                        check_box = check_box_name_buttons(frame_for_table, y1, authorization)
                        list_table_view.append(table_index)
                        list_button_for_table.append(button_for_table)
                        list_height_table_view.append(table_height_for_data_output)
                        list_check_box.append(check_box)
                        button_for_table.clicked.connect(
                            lambda: visible_table_view(list_table_view, list_button_for_table, list_check_box, list_height_table_view, authorization, all_list_button_for_drawing, frame_for_table, list_drawing_button))
                        # перерисовываем кнопки
                        visible_table_view(list_table_view, list_button_for_table, list_check_box, list_height_table_view, authorization, all_list_button_for_drawing, frame_for_table, list_drawing_button)

                # если заполнено одно поле, кроме unit или номера репорта
                if find_data[1] == 2:
                    # делаем запрос в модели
                    if authorization:
                        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
                        cur = conn.cursor()
                        try:
                            if len(cur.execute('''SELECT * FROM {} WHERE "{}" LIKE "%{}%"'''.format(table, find_data[2], find_data[3])).fetchall()) > 0:
                                sqm.setTable(table)
                                sqm.setFilter('"{}" LIKE "%{}%"'.format(find_data[2], find_data[3]))
                                sqm.select()
                        except:
                            continue
                        cur.close()
                    else:
                        sqm.setQuery('''SELECT * FROM {} WHERE "{}" LIKE "%{}%"'''.format(table, find_data[2], find_data[3]))

                    # если не найдено ни одной строчки, то ничего не показываем
                    # список чертежей в рамках одного репорта
                    list_button_for_drawing = []
                    if sqm.rowCount() == 0:
                        table_index.reset()
                    else:
                        list_sqm.append(sqm)
                        # количество строк в найденной таблице
                        count_row_in_table = sqm.rowCount()
                        # сумма всех строк при поиске
                        all_count_row_in_search += count_row_in_table
                        # количество таблиц, в которых найдены данные
                        all_count_table_in_search += 1
                        # высота одной таблицы tableView = количество строк в одной таблице * высоту одной строки + высота строки названия столбцов
                        table_height_for_data_output = count_row_in_table * 25 + 25
                        table_index.hide()
                        y1 += 20
                        table_index.setGeometry(QRect(0, y1, 2500, table_height_for_data_output))
                        # номер линии, сосуда для кнопки названия таблицы
                        line_for_table_name_buttons = sqm.query().value('line')
                        button_for_table = table_name_buttons(frame_for_table, y1, authorization, table, unsort[key_date], language,
                                                              line_for_table_name_buttons)

                        # папка с чертежами для данного репорта
                        dir_with_drawing = table.replace('_', '-')[3:]
                        dir_drawings_db = db.replace('reports', 'drawings')[:-7]
                        # путь, где лежат чертежи для репорта
                        path_drawing = f'{os.path.abspath(os.getcwd())}\\Drawings\\{dir_drawings_db}\\{dir_with_drawing}'
                        # список чертежей для данного репорта
                        try:
                            list_drawing_button = os.listdir(path_drawing)
                        except FileNotFoundError:
                            list_drawing_button = []
                            logger_with_user.info(f'Отсутствуют чертежи для отчёта {dir_with_drawing} в базе чертежей {dir_drawings_db}')
                        # выставляем координату от левого края
                        # если авторизовались
                        if authorization:
                            x11 = 920
                        # если нет
                        else:
                            x11 = 900
                        # итерируемся по количеству чертежей для создания нужного количества кнопок номеров чертежей
                        if len(list_drawing_button) > 0:
                            for index, draw in enumerate(list_drawing_button):
                                # создаём кнопку для чертежа
                                button_for_drawing = drawing_name_buttons(frame_for_table, y1, x11, language, index, path_drawing, draw)
                                x11 += 110
                                list_button_for_drawing.append(button_for_drawing)
                        all_list_button_for_drawing.append(list_button_for_drawing)
                        check_box = check_box_name_buttons(frame_for_table, y1, authorization)
                        list_table_view.append(table_index)
                        list_button_for_table.append(button_for_table)
                        list_height_table_view.append(table_height_for_data_output)
                        list_check_box.append(check_box)
                        button_for_table.clicked.connect(
                            lambda: visible_table_view(list_table_view, list_button_for_table, list_check_box, list_height_table_view, authorization, all_list_button_for_drawing, frame_for_table, list_drawing_button))
                        # перерисовываем кнопки
                        visible_table_view(list_table_view, list_button_for_table, list_check_box, list_height_table_view, authorization, all_list_button_for_drawing, frame_for_table, list_drawing_button)
                # если заполнен номер unit или report_number и любая(-ые) другие данные (номер линии, номер чертежа, номер локации)
                if find_data[1] == 3:
                    # делаем запрос в модели
                    if authorization:
                        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
                        cur = conn.cursor()
                        try:
                            if len(cur.execute('''SELECT * FROM {} WHERE {}'''.format(table, find_data[3])).fetchall()) > 0:
                                sqm.setTable(table)
                                sqm.setFilter(f'{find_data[3]}')
                                sqm.select()
                        except:
                            continue
                        cur.close()
                    else:
                        sqm.setQuery('''SELECT * FROM {} WHERE {}'''.format(table, find_data[3]))
                    # список чертежей в рамках одного репорта
                    list_button_for_drawing = []
                    if sqm.rowCount() == 0:
                        table_index.reset()
                    else:
                        list_sqm.append(sqm)
                        # количество строк в найденной таблице
                        count_row_in_table = sqm.rowCount()
                        # сумма всех строк при поиске
                        all_count_row_in_search += count_row_in_table
                        # количество таблиц, в которых найдены данные
                        all_count_table_in_search += 1
                        # высота одной таблицы tableView = количество строк в одной таблице * высоту одной строки + высота строки названия столбцов
                        table_height_for_data_output = count_row_in_table * 25 + 25
                        table_index.hide()
                        y1 += 20
                        table_index.setGeometry(QRect(0, y1, 2500, table_height_for_data_output))
                        # номер линии, сосуда для кнопки названия таблицы
                        line_for_table_name_buttons = sqm.query().value('line')
                        button_for_table = table_name_buttons(frame_for_table, y1, authorization, table, unsort[key_date], language,
                                                              line_for_table_name_buttons)

                        # папка с чертежами для данного репорта
                        dir_with_drawing = table.replace('_', '-')[3:]
                        dir_drawings_db = db.replace('reports', 'drawings')[:-7]
                        # путь, где лежат чертежи для репорта
                        path_drawing = f'{os.path.abspath(os.getcwd())}\\Drawings\\{dir_drawings_db}\\{dir_with_drawing}'
                        # список чертежей для данного репорта
                        try:
                            list_drawing_button = os.listdir(path_drawing)
                        except FileNotFoundError:
                            list_drawing_button = []
                            logger_with_user.info(f'Отсутствуют чертежи для отчёта {dir_with_drawing} в базе чертежей {dir_drawings_db}')
                        # выставляем координату от левого края
                        # если авторизовались
                        if authorization:
                            x11 = 920
                        # если нет
                        else:
                            x11 = 900
                        # итерируемся по количеству чертежей для создания нужного количества кнопок номеров чертежей
                        if len(list_drawing_button) > 0:
                            for index, draw in enumerate(list_drawing_button):
                                # создаём кнопку для чертежа
                                button_for_drawing = drawing_name_buttons(frame_for_table, y1, x11, language, index, path_drawing, draw)
                                x11 += 110
                                list_button_for_drawing.append(button_for_drawing)
                        all_list_button_for_drawing.append(list_button_for_drawing)
                        check_box = check_box_name_buttons(frame_for_table, y1, authorization)
                        list_table_view.append(table_index)
                        list_button_for_table.append(button_for_table)
                        list_height_table_view.append(table_height_for_data_output)
                        list_check_box.append(check_box)
                        button_for_table.clicked.connect(
                            lambda: visible_table_view(list_table_view, list_button_for_table, list_check_box, list_height_table_view, authorization, all_list_button_for_drawing, frame_for_table, list_drawing_button))
                        # перерисовываем кнопки
                        visible_table_view(list_table_view, list_button_for_table, list_check_box, list_height_table_view, authorization, all_list_button_for_drawing, frame_for_table, list_drawing_button)
                con.close()

    # сообщение о том, что ничего не найдено
    if not list_button_for_table:
        QMessageBox.information(
            window,
            'Упс',
            'Ничего не найдено!'
        )

    # frame_height_for_data_output = all_count_table_in_search * 20
    # высота фрейма = общее количество строк в найденных таблицах * 25 (высота одной строки) + количество таблиц
    # * 2 (кнопка номера репорта и строка названий столбцов) * 20 (высота одной строки) + 20 (высота первой строки с номером первого репорта)
    # + количество таблиц * 20 (расстояние между таблицами в открытом виде)
    # frame_height_for_data_output = all_count_row_in_search * 25 + all_count_table_in_search * 2 * 20 + 20 + all_count_table_in_search * 20
    # frame_for_table.setGeometry(0, 0, 1680, frame_height_for_data_output)
    # размораживаем все кнопки и поля на время поиска данных
    unfreeze_button()
    search_picture.stop_loading()
    window.repaint()
    frame_for_table.show()
    scroll_area.show()


# установка языка RU
def ru():
    button_ru.setCheckable(True)
    button_ru.setChecked(True)
    button_en.setCheckable(False)
    button_kz.setCheckable(False)
    line_search_line.setToolTip('Полный или частичный номер линии/ёмкости, таговый номер')
    line_search_drawing.setToolTip('Полный или частичный номер чертежа')
    line_search_unit.setToolTip('Трёхзначный номер юнита')
    line_search_item_description.setToolTip('Полный(-ое) или частичный(-ое) номер/название локации')
    line_search_number_report.setToolTip('Полный или частичный номер отчёта')
    button_search.setToolTip('Поиск')
    button_print.setToolTip('Печать')
    button_add.setToolTip('Добавить таблицы в БД')
    button_exit.setToolTip('Выход из программы')
    button_ru.setToolTip('Русский язык')
    button_en.setToolTip('Английский язык')
    button_kz.setToolTip('Казахский язык')
    button_faq.setToolTip('Часто задаваемые вопросы')
    line_login.setToolTip('Поле для ввода логина')
    line_password.setToolTip('Поле для ввода пароля')
    button_log_in.setToolTip('Войти в систему')
    button_log_out.setToolTip('Выйти из системы')
    button_delete_table.setToolTip('Удалить выбранную таблицу целиком')
    button_delete_row.setToolTip('Удалить строку из таблицы, на которой находится курсор')
    button_add_row.setToolTip('Добавить строку в конец таблицы')
    button_save.setToolTip('Сохранить изменения')
    button_verification.setToolTip('Автоматическая проверка данных в БД')
    button_statistic_master.setToolTip('Общие данные из БД')
    groupBox_year.setToolTip('Год контроля')
    label_line.setText('Линия/Ёмкость')
    label_drawing.setText('Номер чертежа')
    label_unit.setText('Юнит')
    label_item_description.setText('Номер локации')
    label_number_report.setText('Номер отчёта')
    label_login.setText('Логин')
    label_password.setText('Пароль')
    label_password.setGeometry(QRect(1390, 100, 81, 26))
    button_search.setText('Поиск')
    button_print.setText('Печать')
    button_add.setText('Добавить')
    button_log_in.setText('Войти')
    button_log_out.setText('Выйти')
    button_delete_table.setText('Удалить таблицу')
    button_delete_row.setText('Удалить строку')
    button_add_row.setText('Добавить строку')
    button_save.setText('Сохранить')
    button_statistic_master.setText('Сводные данные')
    button_statistic_master.setGeometry(QRect(1201, 904, 200, 41))
    button_exit.setText('Выход')
    button_verification.setText('Верификация')
    groupBox_location.setTitle('Локация')
    groupBox_ndt.setTitle('Метод контроля')
    groupBox_year.setTitle('Год контроля')
    # переименовываем кнопки найденных репортов
    if window.findChildren(QTableView):
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            for push_button in open_push_button:
                # переименовываем
                old_text = push_button.text()
                if 'number' in old_text:
                    new_text = old_text.replace('report number', 'номер отчёта')
                    new_text = new_text.replace('date', 'дата')
                    new_text = new_text.replace('object of control', 'объект контроля')
                if 'нөмірі' in old_text:
                    new_text = old_text.replace('есеп нөмірі', 'номер отчёта')
                    new_text = new_text.replace('күні', 'дата')
                    new_text = new_text.replace('бақылау объектісі', 'объект контроля')
                if 'drawing' in old_text:
                    new_text = old_text.replace('drawing', 'чертёж')
                if 'сурет' in old_text:
                    new_text = old_text.replace('сурет', 'чертёж')
                push_button.setText(new_text)
                push_button.repaint()
    window.repaint()


# установка языка EN
def en():
    button_ru.setCheckable(False)
    button_en.setCheckable(True)
    button_en.setChecked(True)
    button_kz.setCheckable(False)
    line_search_line.setToolTip('Full or partial line/equipment number, tag number')
    line_search_drawing.setToolTip('Full or partial drawing number')
    line_search_unit.setToolTip('Three-digit unit number')
    line_search_item_description.setToolTip('Full or partial number/name of location')
    line_search_number_report.setToolTip('Full or partial report number')
    button_search.setToolTip('Search')
    button_print.setToolTip('Print')
    button_add.setToolTip('Add tables to the database')
    button_exit.setToolTip('Exit the program')
    button_ru.setToolTip('Russian language')
    button_en.setToolTip('English language')
    button_kz.setToolTip('Kazakh language')
    button_faq.setToolTip('Frequently Asked Questions')
    line_login.setToolTip('Login field')
    line_password.setToolTip('Password field')
    button_log_in.setToolTip('Sign in')
    button_log_out.setToolTip('Sign out')
    button_delete_table.setToolTip('Delete the entire selected table')
    button_delete_row.setToolTip('Delete a row from the table where the cursor is located')
    button_add_row.setToolTip('Add a row to the end of the table')
    button_save.setToolTip('Save changes')
    button_verification.setToolTip('Automatic data verification in the database')
    button_statistic_master.setToolTip('General data from the database')
    groupBox_year.setToolTip('Year of control   ')
    label_line.setText('Line/Equipment')
    label_drawing.setText('Drawing')
    label_unit.setText('Unit')
    label_item_description.setText('Item description')
    label_number_report.setText('Number of report')
    label_login.setText('Login')
    label_password.setText('Password')
    label_password.setGeometry(QRect(1370, 100, 90, 26))
    button_search.setText('Search')
    button_print.setText('Print')
    button_add.setText('Add reports')
    button_log_in.setText('Sign in')
    button_log_out.setText('Sign out')
    button_delete_table.setText('Delete table')
    button_delete_row.setText('Delete row')
    button_add_row.setText('Add row')
    button_save.setText('Save')
    button_statistic_master.setText('Summary data')
    button_statistic_master.setGeometry(QRect(1201, 904, 200, 41))
    button_exit.setText('Exit')
    button_verification.setText('Verification')
    groupBox_location.setTitle('Location')
    groupBox_ndt.setTitle('Method of control')
    groupBox_year.setTitle('Year of control')
    # переименовываем кнопки найденных репортов
    if window.findChildren(QTableView):
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            for push_button in open_push_button:
                # переименовываем
                old_text = push_button.text()
                if 'отчёт' in old_text:
                    new_text = old_text.replace('номер отчёта', 'report number')
                    new_text = new_text.replace('дата', 'date')
                    new_text = new_text.replace('объект контроля', 'object of control')
                if 'нөмірі' in old_text:
                    new_text = old_text.replace('есеп нөмірі', 'report number')
                    new_text = new_text.replace('күні', 'date')
                    new_text = new_text.replace('бақылау объектісі', 'object of control')
                if 'чертёж' in old_text:
                    new_text = old_text.replace('чертёж', 'drawing')
                if 'сурет' in old_text:
                    new_text = old_text.replace('сурет', 'drawing')
                push_button.setText(new_text)
                push_button.repaint()
    window.repaint()


# установка языка KZ
def kz():
    button_ru.setCheckable(False)
    button_en.setCheckable(False)
    button_kz.setCheckable(True)
    button_kz.setChecked(True)
    line_search_line.setToolTip('Толық немесе жартылай жол/сыйымдылық нөмірі, тег нөмірі')
    line_search_drawing.setToolTip('Толық немесе жартылай сызба нөмірі')
    line_search_unit.setToolTip('Үш таңбалы бірлік нөмірі')
    line_search_item_description.setToolTip('Толық немесе ішінара нөмір/орын атауы')
    line_search_number_report.setToolTip('Толық немесе ішінара есеп нөмірі')
    button_search.setToolTip('Іздеу')
    button_print.setToolTip('Мөр')
    button_add.setToolTip('Дерекқорға кестелерді қосыңыз')
    button_exit.setToolTip('Бағдарламадан шығыңыз')
    button_ru.setToolTip('Орыс тілі')
    button_en.setToolTip('Ағылшын тілі')
    button_kz.setToolTip('Қазақ тілі')
    button_faq.setToolTip('Жиі қойылатын сұрақтар')
    line_login.setToolTip('Жүйеге кіру өрісі')
    line_password.setToolTip('Құпия сөз өрісі')
    button_log_in.setToolTip('Кіру')
    button_log_out.setToolTip('Шығу')
    button_delete_table.setToolTip('Таңдалған кестені толығымен жойыңыз')
    button_delete_row.setToolTip('Курсор орналасқан кестеден жолды жойыңыз')
    button_add_row.setToolTip('Кестенің соңына жол қосыңыз')
    button_save.setToolTip('Өзгерістерді сақтау')
    button_verification.setToolTip('Дерекқордағы деректерді автоматты түрде тексеру')
    button_statistic_master.setToolTip('Мәліметтер базасынан жалпы мәліметтер')
    groupBox_year.setToolTip('Бақылау жылы')
    label_line.setText('Сызық/Cыйымд.')
    label_drawing.setText('Сызба нөмірі')
    label_unit.setText('Бірлік')
    label_item_description.setText('Орын нөмірі')
    label_number_report.setText('Есеп нөмірі')
    label_login.setText('Кіру')
    label_password.setText('Құпия сөз')
    label_password.setGeometry(QRect(1370, 100, 100, 26))
    button_search.setText('Іздеу')
    button_print.setText('Мөр')
    button_add.setText('Қосу')
    button_log_in.setText('Кіру үшін')
    button_log_out.setText('Шығу')
    button_delete_table.setText('Кестені жою')
    button_delete_row.setText('Жолды жою')
    button_add_row.setText('Жолды қосыңыз')
    button_save.setText('Сақтау')
    button_statistic_master.setText('Жиынтық деректер')
    button_statistic_master.setGeometry(QRect(1201, 904, 250, 41))
    button_exit.setText('Шығу')
    button_verification.setText('Тексеру')
    groupBox_location.setTitle('Орналасқан жері')
    groupBox_ndt.setTitle('Бақылау әдісі')
    groupBox_year.setTitle('Бақылау жылы')
    # переименовываем кнопки найденных репортов
    if window.findChildren(QTableView):
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            for push_button in open_push_button:
                # переименовываем
                old_text = push_button.text()
                if 'отчёт' in old_text:
                    new_text = old_text.replace('номер отчёта', 'есеп нөмірі')
                    new_text = new_text.replace('дата', 'күні')
                    new_text = new_text.replace('объект контроля', 'бақылау объектісі')
                if 'number' in old_text:
                    new_text = old_text.replace('report number', 'есеп нөмірі')
                    new_text = new_text.replace('date', 'күні')
                    new_text = new_text.replace('object of control', 'бақылау объектісі')
                if 'чертёж' in old_text:
                    new_text = old_text.replace('чертёж', 'сурет')
                if 'drawing' in old_text:
                    new_text = old_text.replace('drawing', 'сурет')
                push_button.setText(new_text)
                push_button.repaint()
    window.repaint()


# выводим данные из таблицы master на лист Excel
def statistic_master():
    output_data_master(define_db_for_search(data_filter_for_search))


# удаляем выбранные репорты
def delete_report():
    scroll_area_with_find_table = window.findChildren(QScrollArea)
    # список индексов выбранных checkbox
    index_check_checkbox = []
    # список таблиц на удаление
    table_for_delete = []
    for area in scroll_area_with_find_table:
        # список всех check_box во Frame
        checkbox = area.findChildren(QCheckBox)
        for index_checkbox, check_checkbox in enumerate(checkbox):
            # если check_box выбран
            if check_checkbox.isChecked():
                # то записываем его индекс
                index_check_checkbox.append(index_checkbox)
        # список всех кнопок во Frame
        all_pushbutton = area.findChildren(QPushButton)
        pushbutton = []
        # выбираем только кнопки с названиями таблиц (исключаем кнопки чертежей)
        for i in all_pushbutton:
            if ':' in i.text():
                pushbutton.append(i)
        for index in index_check_checkbox:
            # список кнопок в которых находится информация о таблице для удаления
            table_for_delete.append(pushbutton[index])
    list_table_and_db_for_delete = del_report(table_for_delete)
    # если не выбрана ни одна таблица для удаления
    if not list_table_and_db_for_delete:
        QMessageBox.information(
            window,
            'Внимание',
            'Вы ничего не выбрали для удаления!'
        )
    # если выбрана хоть одна таблица удаляем её (их)
    else:
        delete_table_from_db(list_table_and_db_for_delete)
        search()


        # # если отображаются найденные таблицы
        # if window.findChildren(QTableView):
        #     open_scroll_area = window.findChildren(QScrollArea)
        #     for scroll in open_scroll_area:
        #         open_check_box = scroll.findChildren(QCheckBox)
        #         # делаем видимые флажки
        #         for check_box in open_check_box:
        #             check_box.show()

# печать репортов
def print_report():
    wbb = openpyxl.Workbook()
    # дата и время формирования файла Excel для печати
    date_time_for_print = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    thin = Side(border_style="thin", color="000000")
    # активатор отсутствия хотя бы одной открытой таблицы
    no_one_open_table = True
    if window.findChildren(QTableView):
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            all_push_button = scroll.findChildren(QPushButton)
            open_push_button = []
            # выбираем только кнопки с названиями таблиц (исключаем кнопки чертежей)
            for i in all_push_button:
                if ':' in i.text():
                    open_push_button.append(i)
            for index_push_button, push_button in enumerate(open_push_button):
                # список найденных заполненных таблиц
                full_sqm = []
                # перебираем найденные таблицы, чтобы исключить нулевые
                for find_table in list_sqm:
                    if find_table.rowCount() != 0:
                        # убираем повторы
                        if find_table not in full_sqm:
                            full_sqm.append(find_table)
                # если таблица открыта, то выводим её на лист Excel
                if push_button.isChecked():
                    no_one_open_table = False
                    # название листа для таблицы
                    name_table = name_table_for_excel_print(push_button.text())
                    # создаём новый лист на каждую таблицу
                    sheet_for_print = wbb.create_sheet(name_table)
                    # вставляем в первую строку название кнопки по выбранной таблицу
                    sheet_for_print.cell(row=1, column=1, value=str(push_button.text()))
                    # выделяем её жирным
                    sheet_for_print.cell(row=1, column=1).font = Font(bold=True)
                    # объединяем в первой строке столбцы 'A:J'
                    sheet_for_print.merge_cells('A1:J1')
                    # вставляем во вторую строку названия столбцов
                    # перебираем названия столбцов
                    for index_column in range(full_sqm[index_push_button].columnCount()):
                        sheet_for_print.cell(row=2, column=index_column + 1,
                                             value=str(full_sqm[index_push_button].query().record().fieldName(index_column)))
                        # выделяем её жирным
                        sheet_for_print.cell(row=2, column=index_column + 1).font = Font(bold=True)
                        # центрируем запись внутри
                        sheet_for_print.cell(row=2, column=index_column + 1).alignment = Alignment(horizontal='center', vertical='center')
                        # заполняем лист Excel
                        for index_row in range(full_sqm[index_push_button].rowCount()):
                            sheet_for_print.cell(row=index_row + 3, column=index_column + 1,
                                                 value=str(full_sqm[index_push_button].record(index_row).value(index_column)))
                            # выделяем основные данные границами
                            sheet_for_print.cell(row=index_row + 3, column=index_column + 1).border = Border(top=thin, left=thin, right=thin,
                                                                                                             bottom=thin)
                        # закрепляем первую строку с названием кнопки, по которой выбрана таблица, и вторую с названиями столбцов
                        sheet_for_print.freeze_panes = "A3"
                        # выделяем её границами
                        sheet_for_print.cell(row=2, column=index_column + 1).border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    # ручной автоподбор ширины столбцов по содержимому
                    ascii_range = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R',
                                   'S', 'T', 'V', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI',
                                   'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AV', 'AX', 'AY', 'AZ']
                    # перебираем все заполненные столбцы
                    for collumn in range(1, full_sqm[index_push_button].columnCount() + 1):
                        max_length_column = 0
                        # перебираем все заполненные строки
                        for roww in range(2, full_sqm[index_push_button].rowCount() + 1):
                            if len(str(sheet_for_print.cell(row=roww, column=collumn).value)) > max_length_column:
                                max_length_column = len(str(sheet_for_print.cell(row=roww, column=collumn).value))
                        # устанавливаем ширину заполненных столбцов по их содержимому
                        sheet_for_print.column_dimensions[ascii_range[collumn]].width = max_length_column + 2
                # путь сохранения в папке с программой
                new_path_for_print = os.path.abspath(os.getcwd()) + '\\Print\\Report for print\\' + date_time_for_print[:7] + '\\'
                if not os.path.exists(new_path_for_print):
                    # то создаём эту папку
                    os.makedirs(new_path_for_print)
                # переменная имени файла с расширением для сохранения и последующего открытия
                name_for_print = str(date_time_for_print) + ' Report for print' + '.xlsx'
            # если ни одна таблица не открыта
            if no_one_open_table:
                QMessageBox.information(
                    window,
                    'Эй',
                    'Вы ничего не выбрали (не открыли) для вывода на печать!'
                )
                return
            # Удаление листа, создаваемого по умолчанию, при создании документа
            del wbb['Sheet']
            # сохраняем файл
            wbb.save(new_path_for_print + name_for_print)
            wbb.close()
            # и открываем его
            os.startfile(new_path_for_print + name_for_print)
            logger_with_user.info('Вывод на печать репорта(ов)\n' + new_path_for_print + name_for_print)


# нажатие на кнопку "FAQ" в главном окне
def faq():
    faq_window = QDialog()
    faq_window.setWindowTitle('Frequently Asked Questions')
    faq_window.setFixedSize(1300, 800)
    if button_ru.isChecked():
        # faq_text_ru - текст внутри на русском языке
        faq_window_text = QLabel(faq_text_ru, faq_window)
    if button_en.isChecked():
        # faq_text_en - текст внутри на английском языке
        faq_window_text = QLabel(faq_text_en, faq_window)
    if button_kz.isChecked():
        # faq_text_kz - текст внутри на казахском языке
        faq_window_text = QLabel(faq_text_kz, faq_window)
    font_faq_window_text = QFont()
    font_faq_window_text.setFamily(u"Arial")
    font_faq_window_text.setPointSize(11)
    faq_window_text.setFont(font_faq_window_text)
    faq_window_text.setGeometry(QRect(0, 0, 1260, 900))
    faq_window_text.setWordWrap(True)

    # область с боковой полосой прокрутки в окне FAQ
    scroll_area_faq = QScrollArea(faq_window)
    scroll_area_faq.setObjectName(u'Scroll_Area_FAQ')
    # полоса прокрутки появляется, только если таблицы больше самой области прокрутки
    scroll_area_faq.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    # задаём размер области с полосой прокрутки
    scroll_area_faq.setGeometry(10, 10, 1280, 700)
    # помещаем frame в область с полосой прокрутки
    scroll_area_faq.setWidget(faq_window_text)
    scroll_area_faq.show()

    # текст кнопки "Понятно" в зависимости от выбранного языка
    if button_ru.isChecked():
        button_close_faq = QPushButton('Понятно', faq_window)
        button_close_faq.setGeometry(600, 730, 100, 50)
    if button_en.isChecked():
        button_close_faq = QPushButton('It\'s clear', faq_window)
        button_close_faq.setGeometry(600, 730, 100, 50)
    if button_kz.isChecked():
        button_close_faq = QPushButton('Ол түсінікті', faq_window)
        button_close_faq.setGeometry(575, 730, 150, 50)
    # присваиваем уникальное объектное имя кнопке "Понятно"
    button_close_faq.setObjectName(u"pushButton_all_clear")
    # задаём параметры стиля и оформления кнопки "Понятно"
    font_button_close_faq = QFont()
    font_button_close_faq.setFamily(u"Arial")
    font_button_close_faq.setPointSize(14)
    button_close_faq.setFont(font_button_close_faq)

    # закрытие окна FAQ
    def close_faq():
        faq_window.accept()

    # нажатие на кнопку "Понятно" для закрытия окна FAQ
    button_close_faq.clicked.connect(close_faq)
    faq_window.exec_()


# создаём пул потока для долгого процесса добавления репортов в БД
threadpool = QThreadPool()


# запускаем добавление новых репортов, замораживаем все окна, поля для ввода и отображаем картинку "Loading"
def start_add_table():
    if window.findChildren(QTableView):
        open_tableview = window.findChildren(QTableView)
        for tableview in open_tableview:
            tableview.setParent(None)
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            for push_button in open_push_button:
                # то сдвигаем вбок кнопки названий таблиц
                push_button.setParent(None)
            open_check_box = scroll.findChildren(QCheckBox)
            for check_box in open_check_box:
                check_box.setParent(None)

    # получаем список файлов отчётов, которые надо обработать
    dir_files = take_dir_files()
    loading = Loading()

    # класс сигнала о завершении работы
    class WorkerSignals(QObject):
        finished = pyqtSignal()

    # разморозка окна приложения по сигналу о завершении процесса добавления репортов в БД в рабочем потоке
    def unfreeze():
        unfreeze_button()
        loading.stop_loading()

    # создаём рабочий поток для добавления репортов
    class Worker(QRunnable):
        def __init__(self, *args, **kwargs):
            super(Worker, self).__init__()
            self.args = args
            self.kwargs = kwargs
            self.signals = WorkerSignals()

        @pyqtSlot()
        def run(self):
            go_add_table(dir_files)
            self.signals.finished.emit()

    worker = Worker()
    worker.signals.finished.connect(unfreeze)
    threadpool.start(worker)
    # замораживаем окно приложения
    freeze_button()
    loading.start_loading()


# класс для отображения и скрытия изображения "Loading"
class Loading():
    def __init__(self):
        super().__init__()

        self.load = QLabel(scroll_area)
        self.load.setGeometry(0, 0, 1660, 630)
        self.load.setAlignment(Qt.AlignCenter)
        self.load.setPixmap(QPixmap(f'{os.path.abspath(os.getcwd())}\\Images\\loading.png'))

    # отображаем "Loading"
    def start_loading(self):
        self.load.show()

    # скрываем "Loading"
    def stop_loading(self):
        self.load.hide()


# класс для отображения и скрытия изображения "Search"
class Searching(Loading):
    def __init__(self):

        self.load = QLabel(scroll_area)
        self.load.setGeometry(0, 0, 1660, 630)
        self.load.setAlignment(Qt.AlignCenter)
        self.load.setPixmap(QPixmap(f'{os.path.abspath(os.getcwd())}\\Images\\searching.png'))

        super().start_loading()
        super().stop_loading()


class Verification(Searching):
    def __init__(self):

        self.load = QLabel(scroll_area)
        self.load.setGeometry(0, 0, 1660, 630)
        self.load.setAlignment(Qt.AlignCenter)
        self.load.setPixmap(QPixmap(f'{os.path.abspath(os.getcwd())}\\Images\\verification.png'))

        super().start_loading()
        super().stop_loading()


# заморозка кнопок и полей для ввода на время загрузки новых репортов
def freeze_button():
    button_search.setDisabled(True)
    button_print.setDisabled(True)
    button_add.setDisabled(True)
    button_delete_table.setDisabled(True)
    button_delete_row.setDisabled(True)
    button_add_row.setDisabled(True)
    button_save.setDisabled(True)
    button_statistic_master.setDisabled(True)
    button_exit.setDisabled(True)
    button_verification.setDisabled(True)
    button_faq.setDisabled(True)
    button_log_in.setDisabled(True)
    button_log_out.setDisabled(True)
    button_ru.setDisabled(True)
    button_en.setDisabled(True)
    button_kz.setDisabled(True)
    line_search_line.setDisabled(True)
    line_search_drawing.setDisabled(True)
    line_search_unit.setDisabled(True)
    line_search_item_description.setDisabled(True)
    line_login.setDisabled(True)
    line_password.setDisabled(True)
    line_search_number_report.setDisabled(True)
    checkBox_on.setDisabled(True)
    checkBox_os.setDisabled(True)
    checkBox_of.setDisabled(True)
    checkBox_utt.setDisabled(True)
    checkBox_paut.setDisabled(True)
    checkBox_2024.setDisabled(True)
    checkBox_2023.setDisabled(True)
    checkBox_2022.setDisabled(True)
    checkBox_2021.setDisabled(True)
    checkBox_2020.setDisabled(True)
    checkBox_2019.setDisabled(True)


# разморозка кнопок и полей для ввода
def unfreeze_button():
    # если пользователь авторизовался
    if authorization:
        button_search.setDisabled(False)
        button_print.setDisabled(False)
        button_add.setDisabled(False)
        button_delete_table.setDisabled(False)
        button_delete_row.setDisabled(False)
        button_add_row.setDisabled(False)
        button_save.setDisabled(False)
        button_statistic_master.setDisabled(False)
        button_exit.setDisabled(False)
        button_verification.setDisabled(False)
        button_faq.setDisabled(False)
        button_log_out.setDisabled(False)
        button_ru.setDisabled(False)
        button_en.setDisabled(False)
        button_kz.setDisabled(False)
        line_search_line.setDisabled(False)
        line_search_drawing.setDisabled(False)
        line_search_unit.setDisabled(False)
        line_search_item_description.setDisabled(False)
        line_login.setDisabled(False)
        line_password.setDisabled(False)
        line_search_number_report.setDisabled(False)
        checkBox_on.setDisabled(False)
        checkBox_os.setDisabled(False)
        checkBox_of.setDisabled(False)
        checkBox_utt.setDisabled(False)
        checkBox_paut.setDisabled(False)
        checkBox_2024.setDisabled(False)
        checkBox_2023.setDisabled(False)
        checkBox_2022.setDisabled(False)
        checkBox_2021.setDisabled(False)
        checkBox_2020.setDisabled(False)
        checkBox_2019.setDisabled(False)
    # если пользователь НЕ авторизовался
    else:
        button_search.setDisabled(False)
        button_print.setDisabled(False)
        button_exit.setDisabled(False)
        button_faq.setDisabled(False)
        button_log_in.setDisabled(False)
        button_ru.setDisabled(False)
        button_en.setDisabled(False)
        button_kz.setDisabled(False)
        line_search_line.setDisabled(False)
        line_search_drawing.setDisabled(False)
        line_search_unit.setDisabled(False)
        line_search_item_description.setDisabled(False)
        line_login.setDisabled(False)
        line_password.setDisabled(False)
        line_search_number_report.setDisabled(False)
        checkBox_on.setDisabled(False)
        checkBox_os.setDisabled(False)
        checkBox_of.setDisabled(False)
        checkBox_utt.setDisabled(False)
        checkBox_paut.setDisabled(False)
        checkBox_2024.setDisabled(False)
        checkBox_2023.setDisabled(False)
        checkBox_2022.setDisabled(False)
        checkBox_2021.setDisabled(False)
        checkBox_2020.setDisabled(False)
        checkBox_2019.setDisabled(False)


# создание окна выбора опций верификаций
def verification_data():
    # закрываем все открытые frame, scroll area, кнопки
    if window.findChildren(QTableView):
        open_tableview = window.findChildren(QTableView)
        for tableview in open_tableview:
            tableview.setParent(None)
        open_scroll_area = window.findChildren(QScrollArea)
        for scroll in open_scroll_area:
            open_push_button = scroll.findChildren(QPushButton)
            for push_button in open_push_button:
                # то сдвигаем вбок кнопки названий таблиц
                push_button.setParent(None)
            open_check_box = scroll.findChildren(QCheckBox)
            for check_box in open_check_box:
                check_box.setParent(None)
    # создаём окно с выбором опций
    ver_window = QDialog()
    ver_window.setWindowTitle('Verification')
    ver_window.setFixedSize(805, 291)

    # создаем checkbox с опциями
    checkbox_all_reports_loading = QCheckBox(ver_window)
    checkbox_all_reports_loading.setObjectName(u"checkbox_checkbox_all_reports_loading")
    checkbox_all_reports_loading.setGeometry(QRect(20, 20, 461, 28))

    # создаем checkbox с опциями
    checkbox_all_tables_loading = QCheckBox(ver_window)
    checkbox_all_tables_loading.setObjectName(u"checkbox_all_tables_loading")
    checkbox_all_tables_loading.setGeometry(QRect(20, 53, 461, 28))

    # создаем checkbox с опциями
    checkbox_duplicate_report = QCheckBox(ver_window)
    checkbox_duplicate_report.setObjectName(u"checkbox_duplicate_report")
    checkbox_duplicate_report.setGeometry(QRect(20, 86, 765, 28))

    # создаем checkbox с опциями
    checkbox_column_in_the_table = QCheckBox(ver_window)
    checkbox_column_in_the_table.setObjectName(u"checkbox_column_in_the_table")
    checkbox_column_in_the_table.setGeometry(QRect(20, 119, 765, 28))

    # создаем checkbox с опциями
    checkbox_drawings_uploaded = QCheckBox(ver_window)
    checkbox_drawings_uploaded.setObjectName(u"checkbox_drawings_uploaded")
    checkbox_drawings_uploaded.setGeometry(QRect(20, 152, 321, 28))

    # создаем checkbox с опциями
    checkbox_unit_column = QCheckBox(ver_window)
    checkbox_unit_column.setObjectName(u"checkbox_unit_column")
    checkbox_unit_column.setGeometry(QRect(20, 185, 765, 28))

    # указываем текст чек-боксов
    if button_ru.isChecked():
        button_start_ver = QPushButton('Начать', ver_window)
        checkbox_all_reports_loading.setText('Все ли отчёты загружены?')
        checkbox_all_tables_loading.setText('Все ли таблицы в отчётах загружены?')
        checkbox_duplicate_report.setText('Есть ли повторяющиеся номера отчётов?')
        checkbox_column_in_the_table.setText('Есть ли в таблицах столбцы "Line" и "Drawing" и все ли они заполнены?')
        checkbox_drawings_uploaded.setText('Все ли чертежи загружены?')
        checkbox_unit_column.setText('Верно ли заполнен столбец "Unit", "Report Date" в сводных данных?')
    if button_en.isChecked():
        button_start_ver = QPushButton('Begin', ver_window)
        checkbox_all_reports_loading.setText('Are all reports uploaded?')
        checkbox_all_tables_loading.setText('Are all tables in reports loaded?')
        checkbox_duplicate_report.setText('Are there duplicate report numbers?')
        checkbox_column_in_the_table.setText('Are there "Line" and "Drawing" columns in the tables and are they all filled?')
        checkbox_drawings_uploaded.setText('Are all drawings uploaded?')
        checkbox_unit_column.setText('Is the "Unit", "Report Date" column filled in correctly in the summary?')
    if button_kz.isChecked():
        button_start_ver = QPushButton('БАСТА', ver_window)
        checkbox_all_reports_loading.setText('Барлық есептер жүктелді ме?')
        checkbox_all_tables_loading.setText('Есептердегі барлық кестелер жүктелді ме?')
        checkbox_duplicate_report.setText('Қайталанатын есеп нөмірлері бар ма?')
        checkbox_column_in_the_table.setText('Кестелерде «Сызық» және «Сызу» бағандары бар ма және олардың барлығы толтырылған ба?')
        checkbox_drawings_uploaded.setText('Барлық сызбалар жүктелді ме?')
        checkbox_unit_column.setText('Қорытындыда «Бірлік», «Есеп беру күні» бағанасы дұрыс толтырылған ба?')

    checkbox_all_reports_loading.setChecked(True)
    checkbox_all_tables_loading.setChecked(True)
    checkbox_duplicate_report.setChecked(True)
    checkbox_column_in_the_table.setChecked(True)
    checkbox_drawings_uploaded.setChecked(True)
    checkbox_unit_column.setChecked(True)

    # задаём положение и форму кнопки начала верификации
    button_start_ver.setGeometry(340, 230, 90, 41)
    # присваиваем уникальное объектное имя кнопке "Начать"
    button_start_ver.setObjectName(u"pushButton_begin")
    # задаём параметры стиля и оформления кнопки "Начать"
    font_button_start_ver = QFont()
    font_button_start_ver.setFamily(u"Arial")
    font_button_start_ver.setPointSize(14)
    button_start_ver.setFont(font_button_start_ver)
    # задаём параметры стиля и оформления check_box 'Все ли отчёты загружены?'
    font_checkbox_all_reports_loading = QFont()
    font_checkbox_all_reports_loading.setFamily(u"Arial")
    font_checkbox_all_reports_loading.setPointSize(13)
    checkbox_all_reports_loading.setFont(font_checkbox_all_reports_loading)
    # задаём параметры стиля и оформления check_box 'Все ли таблицы в репортах загружены?'
    font_checkbox_all_tables_loading = QFont()
    font_checkbox_all_tables_loading.setFamily(u"Arial")
    font_checkbox_all_tables_loading.setPointSize(13)
    checkbox_all_tables_loading.setFont(font_checkbox_all_tables_loading)
    # задаём параметры стиля и оформления check_box 'Есть повторяющиеся номера отчётов?'
    font_checkbox_duplicate_report = QFont()
    font_checkbox_duplicate_report.setFamily(u"Arial")
    font_checkbox_duplicate_report.setPointSize(13)
    checkbox_duplicate_report.setFont(font_checkbox_duplicate_report)
    # задаём параметры стиля и оформления check_box 'Есть ли в таблицах столбцы "Line" и "Drawing"?'
    font_checkbox_column_in_the_table = QFont()
    font_checkbox_column_in_the_table.setFamily(u"Arial")
    font_checkbox_column_in_the_table.setPointSize(13)
    checkbox_column_in_the_table.setFont(font_checkbox_column_in_the_table)
    # задаём параметры стиля и оформления check_box 'Все ли чертежи загружены??'
    font_checkbox_drawings_uploaded = QFont()
    font_checkbox_drawings_uploaded.setFamily(u"Arial")
    font_checkbox_drawings_uploaded.setPointSize(13)
    checkbox_drawings_uploaded.setFont(font_checkbox_drawings_uploaded)
    # задаём параметры стиля и оформления check_box 'Верно ли заполнен столбец "Unit"?'
    font_checkbox_unit_column = QFont()
    font_checkbox_unit_column.setFamily(u"Arial")
    font_checkbox_unit_column.setPointSize(13)
    checkbox_unit_column.setFont(font_checkbox_unit_column)

    # запуск верификации
    def start_ver():
        # создаём файл Excel для вывода результатов верификации
        # name_excel_report_verificarion = create_print_verification()
        # закрываем окно верификации
        ver_window.accept()
        # замораживаем приложение
        freeze_button()
        verificat = Verification()
        verificat.start_loading()
        window.repaint()
        all_reports_loading = False
        if checkbox_all_reports_loading.isChecked():
            all_reports_loading = True
        all_tables_loading = False
        if checkbox_all_tables_loading.isChecked():
            all_tables_loading = True
        duplicate_report = False
        if checkbox_duplicate_report.isChecked():
            duplicate_report = True
        column_in_the_table = False
        if checkbox_column_in_the_table.isChecked():
            column_in_the_table = True
        drawings_uploaded = False
        if checkbox_drawings_uploaded.isChecked():
            drawings_uploaded = True
        unit_column = False
        if checkbox_unit_column.isChecked():
            unit_column = True

        # запускаем саму верификацию
        ver(define_db_for_search(data_filter_for_search), all_reports_loading, all_tables_loading, duplicate_report, column_in_the_table, drawings_uploaded, unit_column)
        unfreeze_button()
        verificat.stop_loading()
        window.repaint()


    # нажатие на кнопку "Начать" для запуска верификации
    button_start_ver.clicked.connect(start_ver)
    ver_window.exec_()


# внесение изменений в БД после нажатия на кнопку "Сохранить"
def save():
    for i in list_sqm:
        for row in range(i.rowCount()):
            for column in range(i.columnCount()):
                if i.isDirty(i.index(row, column)):
                    # имя таблицы в которой надо изменить значение ячейки
                    name_table_update = i.tableName()
                    # номер строки
                    row_number_update = row + 1
                    # имя столбца в котором есть изменения
                    name_column_update = i.record().field(column).name()
                    # значение в ячейке, которое надо записать в БД
                    value_update = i.index(row, column).data()
                    # БД по выбранным фильтрам
                    list_db_update = define_db_for_search(data_filter_for_search)
                    # вносим изменения
                    update_cell(list_db_update, name_table_update, row_number_update, name_column_update, value_update)


# добавление строки в таблицу по нажатию на кнопку "Добавить строку"
def add_row():
    # перебираем области для вывода наёденных данных
    for index, open_table_view in enumerate(list_table_view):
        # если выбрана ячейка в таблице
        if open_table_view.selectionModel().selectedIndexes():
            # вставляем вконец таблицы пустую строку
            list_sqm[index].insertRow(list_sqm[index].rowCount())
            # БД по выбранным фильтрам
            list_db_update = define_db_for_search(data_filter_for_search)
            # имя таблицы в которой надо изменить значение ячейки
            name_table_update = list_sqm[index].tableName()
            # добавляем пустую строку в таблицу в БД
            update_add_row(list_db_update, name_table_update)


# удаление строки из таблицы
def delete_row():
    # перебираем области для вывода наёденных данных
    for index, open_table_view in enumerate(list_table_view):
        # если выбрана ячейка в таблице
        if open_table_view.selectionModel().selectedIndexes():
            # номер строки для удаления
            number_row_delete = open_table_view.selectionModel().selectedIndexes()[0].row() + 1
            # удаляем выделенную строку
            list_sqm[index].removeRow(number_row_delete)
            open_table_view.update()
            # БД по выбранным фильтрам
            list_db_update = define_db_for_search(data_filter_for_search)
            # имя таблицы в которой надо удалить строку
            name_table_update = list_sqm[index].tableName()
            update_delete_row(list_db_update, name_table_update, number_row_delete)


# нажатие кнопки "Войти"
button_log_in.clicked.connect(log_in)

# нажатие на кнопку "Выйти"
button_log_out.clicked.connect(log_out)

# нажатие на кнопку "Поиск"
button_search.clicked.connect(search)
# нажатие на кнопку Enter когда фокус (каретка - мигающий символ "|") находится в поле для ввода номера линии, чертежа, локации или номера репорта
line_search_line.returnPressed.connect(search)
line_search_drawing.returnPressed.connect(search)
line_search_item_description.returnPressed.connect(search)
line_search_number_report.returnPressed.connect(search)
line_search_unit.returnPressed.connect(search)

# нажатие на кнопки "RU", "EN", "KZ"
button_ru.clicked.connect(ru)
button_en.clicked.connect(en)
button_kz.clicked.connect(kz)

# нажатие на кнопку "Сводные данные"
button_statistic_master.clicked.connect(statistic_master)

# нажатие на кнопку "Удалить"
button_delete_table.clicked.connect(delete_report)

# нажатие на кнопку "Удалить строку"
button_delete_row.clicked.connect(delete_row)

# нажатие на кнопку "Добавить строку"
button_add_row.clicked.connect(add_row)

# нажатие на кнопку "Сохранить"
button_save.clicked.connect(save)

# нажатие на кнопку "Печать"
button_print.clicked.connect(print_report)

# нажатие на кнопку "FAQ"
button_faq.clicked.connect(faq)

# нажатие на кнопку "Добавить"
button_add.clicked.connect(start_add_table)


# нажатие на кнопку "Верификация"
button_verification.clicked.connect(verification_data)


def main():
    try:
        # запускаем заставку
        splash_screen()
        # запуск основной программы
        window.show()
        sys.exit(app.exec_())
    finally:
        logger_with_user.info('Программа закрыта\n'
                              '--------------------------------------------------------------------------------')


if __name__ == '__main__':
    main()

# ToDo: Поиск по номеру локации (Location - TML-001, CML-002 DL-01) и описанию (Item_description - Pipe, Elbow, Shell)
# ToDo: При поиске может не быть в таблицах БД столбцов drawing, item_description, location
