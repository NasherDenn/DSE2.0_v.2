import logging
import os
import shutil
import sqlite3
import traceback
import zipfile
import aspose.words as aw
from PyQt5.QtWidgets import *
from YKR.props import *
import PIL
from PIL import Image
import imagehash
import re

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
def write_report_in_db(clear_report: dict, number_report: dict, name_db: str, first_actual_table: list, unit: str, report: str):
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
                                       f'{number_report["report_number"]}\n'
                                       f'{name_table_for_write}\n'
                                       f'{traceback.format_exc()}')
                continue
            for values in clear_report[number_table][1]:
                try:
                    cur.execute('INSERT INTO ' + name_table_for_write + ' VALUES (%s)' % ','.join('?' * len(values)), values)
                    conn.commit()
                    can_write_rep_number_in_master = True
                except sqlite3.OperationalError:
                    logger_with_user.error(f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
                                           f'{values}'
                                           f'{traceback.format_exc()}')
                    continue
        # если таблица удачно записана в БД, то записываем номер репорта, wo, дату в таблицу master
        # number_report - словарь номера репорта, даты, wo
        # first_actual_table - словарь номеров таблиц в которых есть необходимые данные
        # name_table_for_write - имя таблицы
        # name_db - имя БД для записи
        if can_write_rep_number_in_master:
            write_rep_number_in_master(number_report, first_actual_table, name_table_for_write, name_db, unit, report)
    cur.close()


# запись в таблицу master unit, номера репорта, wo, даты, количества таблиц в репорте
def write_rep_number_in_master(number_report: dict, count_table: list, name_table: str, name_db: str, unit: str, report: str):
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
    # если в master нет такого номера репорта, то записываем его первым с значение one_of (1/...) и номером таблицы
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
    # сохраняем чертежи из репорта
    extract_drawing(name_db, report, number_report['report_number'])


# сохраняем чертежи из репорта
def extract_drawing(name_db, report, number_report):
    # определяем название папки для чертежей по номеру репорта
    drawing_number_report = number_report
    name_folder = name_db[:-7].replace('reports', 'drawings')
    # создаём папку для них
    if not os.path.exists(f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}'):
        os.makedirs(f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}')
        # сохраняем в неё фигуры с изображениями из репорта
        doc = aw.Document(report)
        shapes = doc.get_child_nodes(aw.NodeType.SHAPE, True)
        imageIndex = 0
        for shape in shapes:
            shape = shape.as_shape()
            if shape.has_image:
                imageFileName = f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}\\Image.ExportImages.{imageIndex}_{aw.FileFormatUtil.image_type_to_extension(shape.image_data.image_type)}'
                shape.image_data.save(imageFileName)
                imageIndex += 1
        # # сохраняем в неё изображения из репорта
        archive = zipfile.ZipFile(report)
        for file in archive.filelist:
            new_path = f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}'.replace('\\', '/')
            if file.filename.startswith('word/media/'):
                archive.extract(file, path=new_path)
    # удаляем ненужные изображения
    delete_unnecessary_drawing(f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}')


# удаляем ненужные изображения
def delete_unnecessary_drawing(path):
    # перемещаем файлы из word/media/
    if os.path.isdir(f'{path}\\word\\media\\'):
        file_dir = os.listdir(f'{path}\\word\\media\\')
        for file in file_dir:
            shutil.move(f'{path}\\word\\media\\{file}', f'{path}\\{file}')
        # удаляем папки "word" и 'media"
        shutil.rmtree(f'{path}\\word\\media')
        shutil.rmtree(f'{path}\\word')

    # удаляем изображения размером менее 50 кБ
    files_image = os.listdir(f'{path}\\')
    for image in files_image:
        # если файл не PNG, JPEG, BMP
        if not image.endswith('.png') and not image.endswith('.jpg') and not image.endswith('.jpeg') and not image.endswith('.bmp')\
                and not image.endswith('.PNG') and not image.endswith('.JPG') and not image.endswith('.JPEG') and not image.endswith('.BMP'):
            os.remove(f'{path}\\{image}')
            continue
        # если размер файла меньше, чем 50 кБ
        if os.path.getsize(f'{path}\\{image}') < 50000:
            os.remove(f'{path}\\{image}')

    # удаляем стандартные изображения (папка "wrong drawings"), которые не являются чертежами
    new_files_image = os.listdir(f'{path}\\')
    for new_image in new_files_image:
        stop_remove = False
        for wrong_image in os.listdir(f'{os.path.abspath(os.getcwd())}\\Drawings\\wrong drawings'):
            path_image = f'{path}\\{new_image}'
            if not stop_remove:
                if imagehash.dhash(Image.open(path_image)) == imagehash.dhash(
                        Image.open(f'{os.path.abspath(os.getcwd())}\\Drawings\\wrong drawings\\{wrong_image}')):
                    os.remove(f'{path}/{new_image}')
                    stop_remove = True

    # удаляем повторяющиеся изображения
    # список порядковых номеров повторяющихся изображений
    index_image_for_delete = []
    # список оставшихся изображений
    files_image_for_delete = os.listdir(f'{path}\\')
    for index_image in range(len(files_image_for_delete) - 1):
        if imagehash.dhash(Image.open(f'{path}\\{files_image_for_delete[index_image]}')) == imagehash.dhash(
                Image.open(f'{path}\\{files_image_for_delete[index_image + 1]}')):
            index_image_for_delete.append(index_image)
    # перебираем и удаляем повторяющиеся изображения
    for index in index_image_for_delete:
        os.remove(f'{path}\\{files_image_for_delete[index]}')


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


# верификация данных
def ver(list_db: list):
    logger_with_user.info(f'Начало верификации данных\n')
    for db in list_db:
        # подключаемся в базе данных
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
        cur = conn.cursor()
        list_report_number_one_of = cur.execute('''SELECT report_number, one_of FROM master''').fetchall()

        # все ли таблицы в репортах загружены
        for one_of in list_report_number_one_of:
            index_slash = one_of[1].index('/')
            one = one_of[1][:index_slash]
            of = one_of[1][index_slash + 1:]
            if int(one) < int(of):
                logger_with_user.info(f'Не все таблицы {one}/{of} загружены в репорте {one_of[0]}')
            if int(one) > int(of):
                logger_with_user.info(f'{one}/{of} - Такого не может быть {one_of[0]}')

        # уникальны ли номера репортов
        uniq_report_number_one_of = []
        for report_number_one_of in list_report_number_one_of:
            if report_number_one_of[0] not in uniq_report_number_one_of:
                uniq_report_number_one_of.append(report_number_one_of[0])
            else:
                logger_with_user.info(f'Есть повторяющиеся номера репортов {report_number_one_of[0]}!')

        # есть ли в таблице столбцы "Line" и "Drawing" и все ли они заполнены
        # получаем имена всех таблиц в БД
        all_table = cur.execute('''SELECT name FROM main.sqlite_master WHERE type="table"''').fetchall()
        # перебираем их
        for table in all_table:
            line_in_table = ''
            drawing_in_table = ''
            if table[0] != 'master':
                # определяем номер столбцов "Line" и "Drawing
                list_name_column = conn.execute(f'select * from {table[0]}').description
                for row_name_column in list_name_column:
                    if row_name_column[0] == 'Line':
                        line_in_table = cur.execute(f'''SELECT Line FROM {table[0]}''').fetchall()
                    if row_name_column[0] == 'Drawing':
                        drawing_in_table = cur.execute(f'''SELECT Drawing FROM {table[0]}''').fetchall()
                # проверяем столбец "Line"
                if line_in_table:
                    for line_value in line_in_table:
                        # если прочерк
                        if line_value[0] == '-':
                            logger_with_user.info(f'Ошибка в указании номера линии в таблице {table[0]} - знак "-"!')
                            break
                        # если не совпадает с шаблоном
                        if not re.findall('\D\d-\d{3,4}\D?-\D{2}-\d{3}', line_value[0]):
                            logger_with_user.info(f'Ошибка в указании номера линии в таблице {table[0]} - не похож на номер линии или сосуда!')
                            break
                # если в таблице НЕТ столбца "Line"
                else:
                    logger_with_user.info(f'В таблице {table[0]} нет столбца с номером линии!')
                # проверяем столбец "Drawing"
                if drawing_in_table:
                    for drawing_value in drawing_in_table:
                        # если прочерк
                        if drawing_value[0] == '-':
                            logger_with_user.info(f'Ошибка в указании номера чертежа в таблице {table[0]} - знак "-"!')
                            break
                        # если не совпадает с шаблоном
                        if not re.findall('\D{2}\d{2}-\D\d-\d{3,4}\D?-\D{2}', drawing_value[0]):
                            logger_with_user.info(f'Ошибка в указании номера чертежа в таблице {table[0]} - не похож на номер чертежа!')
                            break
                # если в таблице НЕТ столбца "Drawing"
                else:
                    logger_with_user.info(f'В таблице {table[0]} нет столбца с номером чертежа!')
        conn.close()

        # совпадает ли количество папок с чертежами с количеством загруженных репортов
        # получаем список папок с чертежами
        # имя БД с чертежами
        name_folder_drawing = db[:-7].replace('reports', 'drawings')
        # путь к папке с чертежами
        path_dir_drawing = f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder_drawing}\\'
        # список папок с чертежами
        list_dir_drawing = os.listdir(f'{path_dir_drawing}')
        for number_report in list_report_number_one_of:
            # активатор наличия папки с чертежами для загруженного репорта
            drawing_dir_equal_number_report = False
            for number_dir_drawing in list_dir_drawing:
                if number_report[0] == number_dir_drawing:
                    drawing_dir_equal_number_report = True
                    # есть ли в папке с чертежами сами чертежи
                    lll = os.listdir(f'{path_dir_drawing}{number_report[0]}')
                    if len(lll) == 0:
                        logger_with_user.info(f'В БД чертежей {name_folder_drawing} в папке {number_report[0]} отсутствует чертежи!')
                        print(f'В БД чертежей {name_folder_drawing} в папке {number_report[0]} отсутствует чертежи!')
            if not drawing_dir_equal_number_report:
                logger_with_user.info(f'В БД чертежей {name_folder_drawing} отсутствует папка с чертежами для репорта {number_report[0]}!')
                print(f'В БД чертежей {name_folder_drawing} отсутствует папка с чертежами для репорта {number_report[0]}!')

#         # все ли номера репортов в Daily
#         year = ''
#         for i in db:
#             if i.isdigit():
#                 year += i
#         path_master_daily = f'{os.path.abspath(os.getcwd())}\\Master Daily\\Master Daily Activities 20{year}.xlsx'
#         wb = openpyxl.load_workbook(path_master_daily)
#         sheets = wb.sheetnames
#         # список PAUT репортов
#         reports_paut =[]
#         # список UTT репортов
#         reports_utt = []
#         print(db)
#         for month in month_alpha.keys():
#             for month_sheets in sheets:
#                 if month_alpha[month] in month_sheets.upper():
#                     # потом проверка на "дырки" и сопоставление с report_number_one_of[0]
#                     ws = wb[month_sheets]
#                     # перебираем столбец "NDT"
#                     for row_cell in range(2, 2000):
#                         if ws.cell(row=row_cell, column=2).value == 'PAUT':
#                             if ws.cell(row=row_cell, column=3).value[-3:].isdigit():
#                                 reports_paut.append(ws.cell(row=row_cell, column=3).value[-3:])
#                         if ws.cell(row=row_cell, column=2).value == 'UTT':
#                             if ws.cell(row=row_cell, column=3).value[-3:].isdigit():
#                                 reports_utt.append(ws.cell(row=row_cell, column=3).value[-3:])
#         # убираем повторы и сортируем в порядке возрастания номера
#         sort_reports_paut = sorted((list(set(reports_paut))))
#         sort_reports_utt = sorted((list(set(reports_utt))))
#         # список отсутствующих репортов в БД
#         missing_paut_in_db = sort_reports_paut.copy()
#         missing_utt_in_db = sort_reports_utt.copy()
#         print(f'list_report_number_one_of {list_report_number_one_of}')
#         print(f'sort_reports_paut {sort_reports_paut}')
#         print(f'sort_reports_paut {sort_reports_utt}')
#         for report_number_one_of in list_report_number_one_of:
#             for sort_number_reports_paut in sort_reports_paut:
#                 table_paut_in_file = False
#                 # print(f'report_number_one_of[0][-3:] {report_number_one_of[0][-3:]}')
#                 # print(f'sort_number_reports_paut {sort_number_reports_paut}')
#                 if report_number_one_of[0][-3:] in sort_number_reports_paut:
#                     table_paut_in_file = True
#                     table_paut = report_number_one_of[0][-3:]
#             if table_paut_in_file:
#                 missing_paut_in_db.remove(table_paut)
#             for sort_number_reports_utt in sort_reports_utt:
#                 table_utt_in_file = False
#                 if report_number_one_of[0][-3:] not in sort_number_reports_utt:
#                     table_utt_in_file = True
#                     table_utt = report_number_one_of[0][-3:]
#             if table_utt_in_file:
#                 print(table_utt)
#                 missing_utt_in_db.remove(table_utt)
#         print(missing_paut_in_db)
#         print(missing_utt_in_db)
#     pass
