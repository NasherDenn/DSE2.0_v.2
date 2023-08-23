import logging
import os
import sqlite3
import traceback

from PyQt5.QtWidgets import *

# получаем имя машины с которой был осуществлён вход в программу
uname = os.environ.get('USERNAME')
# инициализируем logger
logger = logging.getLogger()
logger_with_user = logging.LoggerAdapter(logger, {'user': uname})


# записываем очищенный репорт в базу данных
# clear_report - очищенные таблицы
# number_report - номер репорта
# name_db - имя БД для записи
# first_actual_table - номер таблицы в репорте для записи
def write_report_in_db(clear_report: dict, number_report: dict, name_db: str, first_actual_table: list, unit: str):
    # меняем все "-" на "_" что бы записать в БД
    true_number_report = number_report['report_number'].replace('-', '_')
    true_number_report = true_number_report.replace('.', '_')
    # подключаемся к БД
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
    conn.commit()
    cur = conn.cursor()
    # создаём таблицу master со столбцами из clear_rep_number, количества таблиц, списка таблиц ЕСЛИ ещё не существует
    if not cur.execute('''SELECT * FROM sqlite_master WHERE type="table" AND name="master"''').fetchall():
        cur.execute('''CREATE TABLE IF NOT EXISTS master (unit, report_number, report_date, work_order, one_of, 
        list_table_report)''')
        conn.commit()
        # создаём индекс
        cur.execute('''CREATE INDEX id ON master (unit, report_number)''')
        conn.commit()
    for number_table in clear_report.keys():
        can_write_rep_number_in_master = False
        # собираем имя таблицы для записи
        name_table_for_write = f'_{number_table}_{true_number_report}'
        if not cur.execute('''SELECT * FROM sqlite_master WHERE  tbl_name="{}"'''.format(name_table_for_write)).fetchone():
            try:
                rep = (",".join(clear_report[number_table][0]))
                cur.execute('''CREATE TABLE IF NOT EXISTS {} ({})'''.format(name_table_for_write, rep))
                conn.commit()
            except sqlite3.OperationalError:
                logger_with_user.error(f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
                                       f'{traceback.format_exc()}')
                continue
            for values in clear_report[number_table][1]:
                try:
                    cur.execute('INSERT INTO ' + name_table_for_write + ' VALUES (%s)' % ','.join('?' * len(values)), values)
                    conn.commit()
                    can_write_rep_number_in_master = True
                except sqlite3.OperationalError:
                    logger_with_user.error(f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
                                           f'{traceback.format_exc()}')
                    continue
        # если таблица удачно записана в БД, то записываем номер репорта, wo, дату в таблицу master
        # number_report - словарь номера репорта, даты, wo
        # first_actual_table - словарь номеров таблиц в которых есть необходимые данные
        # name_table_for_write - имя таблицы
        # name_db - имя БД для записи
        if can_write_rep_number_in_master:
            write_rep_number_in_master(number_report, first_actual_table, name_table_for_write, name_db, unit)
    cur.close()


# запись в таблицу master unit, номера репорта, wo, даты, количества таблиц в репорте
def write_rep_number_in_master(number_report: dict, count_table: list, name_table: str, name_db: str, unit: str):
    # форматируем номер таблицы для лучшей визуализации (меняем "_" на "-")
    name_table = name_table.replace("_", "-")[1:]
    # подключаемся к БД
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
    cur = conn.cursor()
    # создаём таблицу master со столбцами из clear_rep_number, количества таблиц, списка таблиц ЕСЛИ ещё не существует
    if not cur.execute('''SELECT * FROM sqlite_master WHERE type="table" AND name="master"''').fetchall():
        cur.execute('''CREATE TABLE IF NOT EXISTS master (unit, report_number, report_date, work_order, one_of, list_table_report)''')
        conn.commit()
        # создаём индекс
        # conn.commit()
    # если в master нет такого номера репорта, то записываем его первым со значение one_of (1/...) и номером таблицы
    if not cur.execute('''SELECT * FROM master WHERE report_number="{}"'''.format(number_report['report_number'])).fetchone():
        try:
            cur.execute('INSERT INTO master VALUES (?, ?, ?, ?, ?, ?)',
                        (unit,
                         number_report['report_number'],
                         number_report['report_date'],
                         number_report['work_order'],
                         f'1/{len(count_table)}',
                         name_table))
            conn.commit()
        except sqlite3.ProgrammingError:
            logger_with_user.error(f'Ошибка в {number_report["report_number"]}:\n {traceback.format_exc()}')

    # иначе проверяем номер unit
    else:
        # если номер unit такой же, то обновляем строчку с записью в master
        if unit == cur.execute('''SELECT unit FROM master WHERE report_number = "{}"'''.format(number_report['report_number'])).fetchall()[0][0]:
            # пересчитываем и переписываем one_of и дописываем list_table_report номером новой таблицы
            old_one_of = cur.execute('''SELECT one_of FROM master WHERE report_number = "{}"'''.format(number_report['report_number'])).fetchall()[0][
                0]
            index_slash_old_one_of = old_one_of.find('/')
            old_one_of_left_slash = int(old_one_of[:index_slash_old_one_of])
            new_one_of_left_slash = str(old_one_of_left_slash + 1)
            update_one_of = f'{new_one_of_left_slash}{old_one_of[index_slash_old_one_of:]}'
            old_list_table_report = cur.execute('''SELECT list_table_report FROM master WHERE report_number = "{}"'''
                                                .format(number_report['report_number'])).fetchall()[0][0]
            update_list_table_report = f'{old_list_table_report}\n{name_table}'
            cur.execute('''UPDATE  master set one_of='{}', list_table_report='{}' WHERE report_number="{}" AND unit="{}"'''
                        .format(update_one_of,
                                update_list_table_report,
                                number_report['report_number'],
                                unit))
            conn.commit()
        # иначе записываем новую строчку
        else:
            try:
                cur.execute('INSERT INTO master VALUES (?, ?, ?, ?, ?, ?)',
                            (unit,
                             number_report['report_number'],
                             number_report['report_date'],
                             number_report['work_order'],
                             f'1/{len(count_table)}',
                             name_table))
                conn.commit()
            except sqlite3.IntegrityError:
                logger_with_user.error(f'Проверь данные для записи в репорте {number_report}.\n'
                                       f'{traceback.format_exc()}')
    cur.close()


# ищем данные в БД
# db_for_search - список БД, в которых надо искать
# values_for_search - словарь с введёнными значениями для поиска в соответствующее поле (номер линии, номер чертежа, номер локации, номер отчёта)
def look_up_data(db_for_search: list, values_for_search: dict):
    # список таблиц для данной БД для дальнейшего вывода
    find_data = dict()
    # определяем по какому пути будут идти поиски таблиц
    count_value = 0
    for key in values_for_search.keys():
        if values_for_search[key]:
            count_value += 1
    for name_db in db_for_search:
        # подключаемся к БД
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
        cur = conn.cursor()

        # если заполнено только одно поле
        if count_value == 1:
            if values_for_search['unit'] != '' and count_value == 1:
                unit_or_number_report_for_search = values_for_search['unit']
                place_for_search = 'unit'
            if values_for_search['number_report'] != '' and count_value == 1:
                unit_or_number_report_for_search = values_for_search['number_report']
                place_for_search = 'report_number'

            # если заполнено только поле unit или number_report
            if values_for_search['number_report'] != '' or values_for_search['unit'] != '':
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute('''SELECT "list_table_report", "report_date" FROM master WHERE "{}" LIKE "%{}%"'''.
                                                                   format(place_for_search, unit_or_number_report_for_search)).fetchall()
                # преобразуем найденные названия таблиц в вид, в котором они записаны в БД
                list_table = transform_name_table(find_tables_by_unit_or_report_number)
                find_data[name_db] = list_table
                place_for_search = ''
                values = ''
                path = 1

            # если заполнено одно поле, кроме unit или номера репорта
            else:
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute('''SELECT "list_table_report", "report_date" FROM master''').fetchall()
                # преобразуем найденные названия таблиц в вид, в котором они записаны в БД
                list_table = transform_name_table(find_tables_by_unit_or_report_number)
                find_data[name_db] = list_table
                # определяем какое поле заполнено
                if values_for_search['line'] != '':
                    place_for_search = 'line'
                    values = values_for_search['line']
                if values_for_search['drawing'] != '':
                    place_for_search = 'drawing'
                    values = values_for_search['drawing']
                if values_for_search['item_description'] != '':
                    place_for_search = 'item_description'
                    values = values_for_search['item_description']
                path = 2

        # если заполнено больше, чем одно поле
        if count_value > 1:
            # если заполнены только номер unit и номер репорта
            if values_for_search['unit'] != '' and values_for_search['number_report'] != '' and count_value == 2:
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute(
                    '''SELECT "list_table_report", "report_date" FROM master WHERE "unit" LIKE "%{}%" and "report_number" LIKE "%{}%"'''.
                    # '''SELECT "list_table_report", "report_date" FROM master WHERE "unit"="{}" and "report_number"="{}"'''.
                    format(values_for_search['unit'], values_for_search['number_report'])).fetchall()
                # преобразуем найденные названия таблиц в вид, в котором они записаны в БД
                list_table = transform_name_table(find_tables_by_unit_or_report_number)
                find_data[name_db] = list_table
                place_for_search = ''
                values = ''
                path = 1

            # если заполнен номер unit или report_number и любая(-ые) другие данные (номер линии, номер чертежа, номер локации)
            elif (values_for_search['unit'] != '' or values_for_search['number_report'] != '') and count_value > 1:
                if values_for_search['unit'] != '':
                    unit_or_number_report_for_search = values_for_search['unit']
                    place_for_search = 'unit'
                if values_for_search['number_report'] != '':
                    unit_or_number_report_for_search = values_for_search['number_report']
                    place_for_search = 'report_number'
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute(
                    '''SELECT "list_table_report", "report_date" FROM master WHERE "{}"="{}"'''.format(place_for_search,
                                                                                                       unit_or_number_report_for_search)). \
                    fetchall()
                # преобразуем найденные названия таблиц в вид, в котором они записаны в БД
                list_table = transform_name_table(find_tables_by_unit_or_report_number)
                find_data[name_db] = list_table
                place_for_search = ''
                # формируем условия для конечного поиска, в зависимости от количества полей и какие именно поля заполнены
                comma_or_and = "comma"
                values = conversion_to_accumulation(values_for_search, comma_or_and)[:-2]
                path = 3

            # если заполнены любые поля (номер линии, номер чертежа, номер локации), КРОМЕ номера unit и номера репорта
            if (values_for_search['unit'] == '' and values_for_search['number_report'] == '') and count_value > 1:
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute('''SELECT "list_table_report", "report_date" FROM master''').fetchall()
                # преобразуем найденные названия таблиц в вид, в котором они записаны в БД
                list_table = transform_name_table(find_tables_by_unit_or_report_number)
                find_data[name_db] = list_table
                place_for_search = ''
                # формируем условия для конечного поиска, в зависимости от количества полей и какие именно поля заполнены
                comma_or_and = 'and'
                values = conversion_to_accumulation(values_for_search, comma_or_and)[:-5]
                path = 3
        cur.close()
    return find_data, path, place_for_search, values


# преобразуем найденные названия таблиц в вид, в котором они записаны в БД
def transform_name_table(name_table: list) -> dict:
    # итоговый словарь отсортированных номеров таблиц их дат
    new_list_table_and_date = {}
    # индикатор первого прохода поиска
    first_check_date = True
    for i in name_table:
        # если первый проход
        if first_check_date:
            new_list_table = remove_add_change(i[0])
            new_list_table_and_date[i[1]] = sorted(new_list_table)
            first_check_date = False
        # иначе начинаем проверять все возможные комбинации новой даты и предыдущей
        else:
            # если новая дата уже есть в конечном словаре
            if i[1] in new_list_table_and_date.keys():
                new_list_table = remove_add_change(i[0])
                # перезаписываем номера таблиц в дате, которая уже есть в конечном словаре
                new_list_table_and_date[i[1]] = new_list_table_and_date[i[1]] + sorted(new_list_table)
            # иначе, если новой даты нет в конечном словаре
            else:
                new_list_table = remove_add_change(i[0])
                # записываем новые номера таблиц с новой датой в конечный словарь
                new_list_table_and_date[i[1]] = sorted(new_list_table)
    return new_list_table_and_date


# обрабатываем номера таблиц - убираем повторы, меняем и добавляем знаки '-', '_'
def remove_add_change(raw_tables: str) -> list:
    # обнуляем список обработанных номеров таблиц
    new_list_table = []
    list_table = raw_tables.split('\n')
    list_table = list(set(list_table))
    for ii in list_table:
        name_table = '_' + ii.replace('-', '_')
        new_list_table.append(name_table)
    return new_list_table


# формируем условия для конечного поиска, в зависимости от количества полей и какие именно поля заполнены
def conversion_to_accumulation(values_for_search: dict, comma_or_and: str) -> str:
    # определяем поля и данные из values_for_search для конечного поиска
    completed_keys = []
    completed_values = []
    for key in values_for_search.keys():
        if values_for_search[key] != '':
            if key != 'unit' and key != 'number_report':
                completed_keys.append(key)
                completed_values.append(values_for_search[key])
    count_completed_keys = len(completed_keys)
    accumulation_variable = ''
    # формируем условия для конечного поиска, в зависимости от количества полей и какие именно поля заполнены
    for i in range(count_completed_keys):
        added_search_data = f'{completed_keys[i].capitalize()} LIKE "%{completed_values[i]}%"'
        if comma_or_and == 'comma':
            accumulation_variable += added_search_data + ', '
        if comma_or_and == 'and':
            accumulation_variable += added_search_data + ' and '
    return accumulation_variable


# получаем данные из таблицы master
def data_master(db):
    # подключаемся в базе данных
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
    cur = conn.cursor()
    # получаем все данные (кроме столбца unit) из таблицы master и сортируем по номеру репорта
    myself_data = cur.execute("SELECT * FROM master ORDER BY report_number DESC").fetchall()
    cur.close()
    return myself_data


# удаляем выбранные таблицы из соответствующих БД
def delete_table_from_db(list_table_and_db_for_delete):
    # спрашиваем, точно ли надо удалять репорт(ы)
    question_delete = QMessageBox()
    question_delete.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    question_delete.setWindowTitle('Внимание')
    question_delete.setText('Вы уверены, что хотите удалить данные отчёты?')
    # если нажата кнопка "Да", то
    if question_delete.exec() == QMessageBox.Yes:
        for db in list_table_and_db_for_delete.keys():
            # подключаемся к БД
            conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
            cur = conn.cursor()
            for table in list_table_and_db_for_delete[db]:
                cur.execute('DROP TABLE {}'.format(table))
                conn.commit()
                logger_with_user.warning(f'БЫЛА УДАЛЕНА ТАБЛИЦА {table}')
                # обновляем данные в таблице master
                # если в sqlite_master не осталось таблиц из этого репорта, то удаляем номер репорта из master
                if not cur.execute('SELECT * FROM sqlite_master WHERE  name LIKE "%{}%"'.format(f'{table}'[2:])).fetchall():
                    cur.execute('DELETE from master WHERE report_number="{}"'.format(f'{table.replace("_", "-")}'[3:]))
                    conn.commit()
                # иначе обновляем данные в таблице master - уменьшаем на единицу one_of и удаляем строчку номера репорта из list_table_report
                else:
                    update_master_by_delete(table, db)
            cur.close()
    else:
        return


# обновляем данные в таблице master - уменьшаем на единицу one_of и удаляем строчку из list_table_report
def update_master_by_delete(table, db):
    table = table.replace('_', '-')[1:]
    # подключаемся в базе данных
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
    cur = conn.cursor()
    # получаем номера всех записанных таблиц в виде строки
    variable_report_for_delete_from_master = cur.execute(
        'SELECT list_table_report FROM master WHERE list_table_report LIKE "%{}%"'.format(table)).fetchall()
    # определяем номер репорта для удаления
    variable = cur.execute('SELECT report_number FROM master WHERE list_table_report LIKE "%{}%"'.format(table)).fetchall()[0][0]
    # удаляем в найденной строке, с номерами всех записанных таблиц, выбранную таблицу
    variable_report_for_delete_from_master_intermediate = variable_report_for_delete_from_master[0][0].split('\n')
    variable_report_for_delete_from_master_intermediate.remove(table)
    variable_report_for_delete_from_master_new = '\n'.join(variable_report_for_delete_from_master_intermediate)
    # получаем количество репортов в столбце 'one_of'
    one_of_column = cur.execute('SELECT one_of FROM master WHERE list_table_report LIKE "%{}%"'.format(table)).fetchall()[0][0]
    # получаем номер позиции символа '/'
    index_one_of_load = one_of_column.index('/')
    # уменьшаем на 1 количество записанных таблиц в столбце 'one_of'
    one_of_load_new = str(int(one_of_column[:index_one_of_load]) - 1)
    # определяем новое значение 'one_of'
    one_of_new = one_of_load_new + str(one_of_column[index_one_of_load:])
    # обновляем ячейку с количеством записанных таблиц
    cur.execute('UPDATE master SET one_of = "{}" WHERE report_number = "{}"'.format(one_of_new, variable))
    # обновляем ячейку с номерами всех оставшихся таблиц
    conn.commit()
    cur.execute('UPDATE master SET list_table_report = "{}" WHERE report_number = "{}"'.format(
        variable_report_for_delete_from_master_new, variable))
    conn.commit()
    cur.close()
