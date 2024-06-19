import logging
# import os
import shutil
import sqlite3
import traceback
import zipfile
import aspose.words as aw
import openpyxl
from PyQt5.QtWidgets import *

import YKR.utilities_interface
from YKR.props import *
from YKR.utilities_interface import *
import PIL
from PIL import Image
import imagehash
import re
# from collections import defaultdict
import numpy as np
import copy

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
def write_report_in_db(clear_report: dict, number_report: dict, name_db: str, first_actual_table: list, unit: str,
                       report: str):
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
        if not cur.execute(
                '''SELECT * FROM sqlite_master WHERE  tbl_name="{}"'''.format(name_table_for_write)).fetchone():
            try:
                rep = (",".join(clear_report[number_table][0]))
                cur.execute('''CREATE TABLE IF NOT EXISTS {} ({})'''.format(name_table_for_write, rep))
                conn.commit()
            except sqlite3.OperationalError:
                logger_with_user.error(
                    f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
                    f'{name_table_for_write}\n'
                    f'{rep}\n'
                    f'{traceback.format_exc()}')
                continue
            for values in clear_report[number_table][1]:
                try:
                    cur.execute('INSERT INTO ' + name_table_for_write + ' VALUES (%s)' % ','.join('?' * len(values)),
                                values)
                    conn.commit()
                    can_write_rep_number_in_master = True
                except sqlite3.OperationalError:
                    logger_with_user.error(
                        f'В репорте {number_report["report_number"]} таблице {name_table_for_write} какая-то ошибка! А именно:\n'
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
def write_rep_number_in_master(number_report: dict, count_table: list, name_table: str, name_db: str, unit: str,
                               report: str):
    # форматируем номер таблицы для лучшей визуализации (меняем "_" на "-")
    name_table = name_table.replace("_", "-")[1:]
    # подключаемся к БД
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
    cur = conn.cursor()
    # создаём таблицу master со столбцами из clear_rep_number, количества таблиц, списка таблиц ЕСЛИ ещё не существует
    if not cur.execute('''SELECT * FROM sqlite_master WHERE type="table" AND name="master"''').fetchall():
        cur.execute(
            '''CREATE TABLE IF NOT EXISTS master (unit, report_number, report_date, work_order, one_of, list_table_report)''')
        conn.commit()
    # если в master нет такого номера репорта, то записываем его первым со значением one_of (1/...) и номером таблицы
    if not cur.execute(
            '''SELECT * FROM master WHERE report_number="{}"'''.format(number_report['report_number'])).fetchone():
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
        if unit == cur.execute('''SELECT unit FROM master WHERE report_number = "{}"'''.format(
                number_report['report_number'])).fetchall()[0][0]:
            # пересчитываем и переписываем one_of и дописываем list_table_report номером новой таблицы
            old_one_of = cur.execute('''SELECT one_of FROM master WHERE report_number = "{}"'''.format(
                number_report['report_number'])).fetchall()[0][
                0]
            index_slash_old_one_of = old_one_of.find('/')
            old_one_of_left_slash = int(old_one_of[:index_slash_old_one_of])
            new_one_of_left_slash = str(old_one_of_left_slash + 1)
            update_one_of = f'{new_one_of_left_slash}{old_one_of[index_slash_old_one_of:]}'
            old_list_table_report = cur.execute('''SELECT list_table_report FROM master WHERE report_number = "{}"'''
                                                .format(number_report['report_number'])).fetchall()[0][0]
            update_list_table_report = f'{old_list_table_report}\n{name_table}'
            cur.execute(
                '''UPDATE  master set one_of='{}', list_table_report='{}' WHERE report_number="{}" AND unit="{}"'''
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


# запись в таблицу master отчёты по сварке и поперечном сканировании
def write_in_master_nonreports(report_number: str, name_db: str, welding: bool, shaer_wave: bool):
    # подключаемся к БД
    conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
    cur = conn.cursor()
    # если в master нет такого номера репорта, то записываем его первым со значением one_of (1/...) и номером таблицы
    if not cur.execute('''SELECT * FROM master WHERE report_number="{}"'''.format(report_number)).fetchone():
        if welding:
            msg = 'WELDING'
        if shaer_wave:
            msg = 'SHAER WAVE'
        try:
            cur.execute('INSERT INTO master VALUES (?, ?, ?, ?, ?, ?)',
                        ('-',
                         report_number,
                         '-',
                         '-',
                         '-',
                         msg))
            conn.commit()
        except:
            logger_with_user.error(f'Ошибка в {report_number}:\n {traceback.format_exc()}')
    cur.close()


# извлекаем и сохраняем чертежи из репорта
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
        image_index = 0
        for shape in shapes:
            shape = shape.as_shape()
            if shape.has_image:
                image_file_name = f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}\\Image.ExportImages.{image_index}_{aw.FileFormatUtil.image_type_to_extension(shape.image_data.image_type)}'
                shape.image_data.save(image_file_name)
                image_index += 1
        # сохраняем в неё изображения из репорта
        archive = zipfile.ZipFile(report)
        for file in archive.filelist:
            new_path = f'{os.path.abspath(os.getcwd())}\\Drawings\\{name_folder}\\{drawing_number_report}'.replace('\\',
                                                                                                                   '/')
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

    # удаляем изображения размером менее 50 кБ и файлы, которые не являются изображениями
    files_image = os.listdir(f'{path}\\')
    for image in files_image:
        # если файл не PNG, JPEG, BMP
        if not image.endswith('.png') and not image.endswith('.jpg') and not image.endswith(
                '.jpeg') and not image.endswith('.bmp') \
                and not image.endswith('.PNG') and not image.endswith('.JPG') and not image.endswith(
            '.JPEG') and not image.endswith('.BMP'):
            os.remove(f'{path}\\{image}')
            continue
        # # если размер файла меньше, чем 50 кБ
        # if os.path.getsize(f'{path}\\{image}') < 50000:
        #     os.remove(f'{path}\\{image}')

    # путь к проверяемой папке с номером репорта, где хранятся изображения
    # remaining_files_image = os.listdir(f'{path}\\')
    # удаляем дубликаты
    remove_duplicate = DuplicateRemover(path)
    remove_duplicate.find_duplicates()
    # удаляем похожие неверные изображения
    for wrong_image in os.listdir(f'{os.path.abspath(os.getcwd())}\\Drawings\\wrong drawings'):
        remove_duplicate.find_similar(f'{os.path.abspath(os.getcwd())}\\Drawings\\wrong drawings\\{wrong_image}')


# класс для удаления дубликатов изображений
class DuplicateRemover:
    def __init__(self, dirname, hash_size=8):
        self.dirname = dirname
        self.hash_size = hash_size

    def find_duplicates(self):
        file_names = os.listdir(self.dirname)
        hashes = {}
        duplicates = []
        for image in file_names:
            try:
                with Image.open(os.path.join(self.dirname, image)) as img:
                    temp_hash = imagehash.average_hash(img, self.hash_size)
                    if temp_hash in hashes:
                        duplicates.append(image)
                    else:
                        hashes[temp_hash] = image
            except:
                continue
        if len(duplicates) != 0:
            for duplicate in duplicates:
                os.remove(os.path.join(self.dirname, duplicate))

    def find_similar(self, location, similarity=80):
        file_names = os.listdir(self.dirname)
        threshold = 1 - similarity / 100
        diff_limit = int(threshold * (self.hash_size ** 2))

        with Image.open(location) as img:
            hash1 = imagehash.average_hash(img, self.hash_size).hash

        for image in file_names:
            try:
                with Image.open(os.path.join(self.dirname, image)) as img:
                    hash2 = imagehash.average_hash(img, self.hash_size).hash

                    if np.count_nonzero(hash1 != hash2) <= diff_limit:
                        os.remove(os.path.join(self.dirname, image))
            except:
                continue


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
            # заполнено только поле unit или number_report
            if values_for_search['unit'] != '' and count_value == 1:
                unit_or_number_report_for_search = values_for_search['unit']
                place_for_search = 'unit'
            if values_for_search['number_report'] != '' and count_value == 1:
                unit_or_number_report_for_search = values_for_search['number_report']
                place_for_search = 'report_number'

            # поиск, если заполнено только поле unit или number_report
            if (values_for_search['number_report'] != '' or values_for_search['unit'] != '') and count_value == 1:
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute(
                    '''SELECT "list_table_report", "report_date" FROM master WHERE "{}" LIKE "%{}%"'''.
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
                find_tables_by_unit_or_report_number = cur.execute(
                    '''SELECT "list_table_report", "report_date" FROM master''').fetchall()
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
                if values_for_search['location'] != '':
                    place_for_search = 'location'
                    values = values_for_search['location']
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
                    '''SELECT "list_table_report", "report_date" FROM master WHERE "{}" LIKE "%{}%"'''.format(
                        place_for_search,
                        unit_or_number_report_for_search)). \
                    fetchall()
                # преобразуем найденные названия таблиц в вид, в котором они записаны в БД
                list_table = transform_name_table(find_tables_by_unit_or_report_number)
                find_data[name_db] = list_table
                place_for_search = ''
                # формируем условия для конечного поиска, в зависимости от количества полей и какие именно поля заполнены
                # comma_or_and = "comma"
                comma_or_and = "and"
                # values = conversion_to_accumulation(values_for_search, comma_or_and)[:-2]
                values = conversion_to_accumulation(values_for_search, comma_or_and)[:-5]
                path = 3

            # если заполнены любые поля (номер линии, номер чертежа, номер локации), КРОМЕ номера unit и номера репорта
            if (values_for_search['unit'] == '' and values_for_search['number_report'] == '') and count_value > 1:
                # находим названия таблиц
                find_tables_by_unit_or_report_number = cur.execute(
                    '''SELECT "list_table_report", "report_date" FROM master''').fetchall()
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
                if not cur.execute(
                        'SELECT * FROM sqlite_master WHERE  name LIKE "%{}%"'.format(f'{table}'[2:])).fetchall():
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
    variable = \
    cur.execute('SELECT report_number FROM master WHERE list_table_report LIKE "%{}%"'.format(table)).fetchall()[0][0]
    # удаляем в найденной строке, с номерами всех записанных таблиц, выбранную таблицу
    variable_report_for_delete_from_master_intermediate = variable_report_for_delete_from_master[0][0].split('\n')
    variable_report_for_delete_from_master_intermediate.remove(table)
    variable_report_for_delete_from_master_new = '\n'.join(variable_report_for_delete_from_master_intermediate)
    # получаем количество репортов в столбце 'one_of'
    one_of_column = \
    cur.execute('SELECT one_of FROM master WHERE list_table_report LIKE "%{}%"'.format(table)).fetchall()[0][0]
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
def ver(list_db: list, all_reports_loading, all_tables_loading, duplicate_report, column_in_the_table,
        drawings_uploaded, unit_column):
    logger_with_user.info(f'Начало верификации данных\n')
    for db in list_db:
        # подключаемся в базе данных
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{db}')
        cur = conn.cursor()
        list_report_number_one_of = cur.execute(
            '''SELECT report_number, one_of, unit, report_date, list_table_report FROM master''').fetchall()

        # все ли таблицы в репортах загружены
        if all_tables_loading:
            for one_of in list_report_number_one_of:
                if 'WELDING' not in one_of[4] and 'SHAER WAVE' not in one_of[4]:
                    index_slash = one_of[1].index('/')
                    one = one_of[1][:index_slash]
                    of = one_of[1][index_slash + 1:]
                    if int(one) < int(of):
                        logger_with_user.info(f'Не все таблицы {one}/{of} загружены в репорте {one_of[0]}')

                        # YKR.utilities_interface.print_verification(name_excel_report_verificarion[0], name_excel_report_verificarion[1], 'Количество загруженный таблиц', one, of, one_of[0])

                    if int(one) > int(of):
                        logger_with_user.info(f'{one}/{of} - Такого не может быть {one_of[0]}')

                        # YKR.utilities_interface.print_verification(name_excel_report_verificarion[0], name_excel_report_verificarion[1], 'Количество загруженный таблиц', one, of, one_of[0])

        # уникальны ли номера репортов
        if duplicate_report:
            uniq_report_number_one_of = []
            for report_number_one_of in list_report_number_one_of:
                if report_number_one_of[0] not in uniq_report_number_one_of:
                    uniq_report_number_one_of.append(report_number_one_of[0])
                else:
                    logger_with_user.info(f'Есть повторяющиеся номера репортов {report_number_one_of[0]}!')

        # есть ли в таблице столбцы "Line" и "Drawing" и все ли они заполнены
        if column_in_the_table:
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
                            if not line_in_table:
                                logger_with_user.warning(f'В таблице {table[0]} отсутствуют данные!')
                        if row_name_column[0] == 'Drawing':
                            drawing_in_table = cur.execute(f'''SELECT Drawing FROM {table[0]}''').fetchall()
                    # проверяем столбец "Line"
                    if line_in_table:
                        for line_value in line_in_table:
                            # если прочерк
                            if line_value[0] == '-':
                                logger_with_user.warning(
                                    f'Ошибка в указании номера линии в таблице {table[0]} - знак "-"!')
                                break
                            # если не совпадает с шаблоном или
                            if not re.findall('\D\d-\d{3,4}\D?-\D{2}-(\d{3}|\D{2})', line_value[0]):
                                if 'FRACK' not in line_value[0].upper():
                                    if 'TANK' not in line_value[0].upper():
                                        logger_with_user.warning(
                                            f'Ошибка в указании номера линии ({line_value[0]}) в таблице {table[0]} - не похож '
                                            f'на номер линии или сосуда или пропущены буквы/цифры!')
                                        break

                    # если в таблице НЕТ столбца "Line"
                    else:
                        logger_with_user.warning(f'В таблице {table[0]} нет столбца с номером линии!')
                    # проверяем столбец "Drawing"
                    if drawing_in_table:
                        for drawing_value in drawing_in_table:
                            # если прочерк
                            if drawing_value[0] == '-':
                                logger_with_user.warning(
                                    f'Ошибка в указании номера чертежа в таблице {table[0]} - знак "-"!')
                                break
                            # если не совпадает с шаблоном
                            if not re.findall('\D{2}\d{2}-\D\d-\d{3,4}\D?-\D{2}', drawing_value[0]):
                                logger_with_user.warning(
                                    f'Ошибка в указании номера чертежа ({drawing_value[0]}) в таблице {table[0]} - не похож '
                                    f'на номер чертежа или пропущены буквы/цифры!')
                                break

        # совпадает ли количество папок с чертежами с количеством загруженных репортов
        if drawings_uploaded:
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
                            logger_with_user.info(
                                f'В БД чертежей {name_folder_drawing} в папке {number_report[0]} отсутствуют чертежи!')
                if not drawing_dir_equal_number_report:
                    logger_with_user.info(
                        f'В БД чертежей {name_folder_drawing} отсутствует папка с чертежами для репорта {number_report[0]}!')

        # Верно ли заполнен столбец "Unit", "Report Date" в таблице master
        if unit_column:
            for index_unit, unit in enumerate(list_report_number_one_of):
                # столбец "Unit"
                if unit[2] == '-':
                    if 'WELDING' not in unit[4] and 'SHAER WAVE' not in unit[4]:
                        logger_with_user.warning(
                            f'Не указан номер юнита ({unit[2]}) в отчёте {unit[0]} в сводных данных БД {db}!')
                elif len(unit[2]) < 3:
                    logger_with_user.warning(
                        f'Не полный номер юнита ({unit[2]}) в отчёте {unit[0]} в сводных данных БД {db}!')
                elif 'FRACK' not in unit[2].upper():
                    for letter in unit[2]:
                        if not letter.isdigit():
                            logger_with_user.warning(
                                f'В номере юнита ({unit[2]}) не должно быть букв, отчёт {unit[0]} в сводных данных БД {db}!')
                            break
                # столбец "Report Date"
                if not re.findall('\d{2}(-\d{2})?\.\d{2}\.\d{4}', unit[3]):
                    if 'WELDING' not in unit[4] and 'SHAER WAVE' not in unit[4]:
                        logger_with_user.warning(
                            f'Дата ({unit[3]}) отчёта {unit[0]} в сводных данных БД {db} не соответствует шаблону '
                            f'(ХХ.ХХ.ХХХХ или ХХ-ХХ.ХХ.ХХХХ)!')

        # все ли репорты загружены
        if all_reports_loading:
            if len(list_report_number_one_of) != count_reports_in_ndt[db]:
                miss_reports = count_reports_in_ndt[db] - len(list_report_number_one_of)
                if miss_reports > 0:
                    logger_with_user.info(f'В БД {db} не загружено {miss_reports} отчётов!')
                if miss_reports < 0:
                    logger_with_user.info(f'В БД {db} есть лишние ({miss_reports}) загруженные отчёты.')

        conn.close()
    logger_with_user.info(f'Верификации данных завершена\n')


# обновление значения ячейки в БД при нажатии на кнопку "Сохранить"
def update_cell(list_db_update, name_table_update, row_number_update, name_column_update, value_update):
    for name_db in list_db_update:
        # подключаемся к БД
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
        cur = conn.cursor()
        # определяем  БД
        if cur.execute('''SELECT name FROM sqlite_master WHERE type='table' AND name="{}"'''.format(
                name_table_update)).fetchone():
            # получаем ROWID по номеру строки (row_number_update)
            list_row_id = []
            for row_id in cur.execute('''SELECT ROWID FROM {}'''.format(name_table_update)).fetchall():
                list_row_id.append(row_id[0])
            # вносим изменения
            cur.execute(
                '''UPDATE {} SET {}="{}" WHERE ROWID="{}"'''.format(name_table_update, name_column_update, value_update,
                                                                    list_row_id[row_number_update - 1]))
            conn.commit()
        cur.close()


# добавление пустой строки в таблицу
def update_add_row(list_db_update, name_table_update):
    for name_db in list_db_update:
        # подключаемся к БД
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
        cur = conn.cursor()
        # определяем  БД
        if cur.execute('''SELECT name FROM sqlite_master WHERE type='table' AND name="{}"'''.format(
                name_table_update)).fetchone():
            # кортеж названий столбцов в таблице
            list_name_column = conn.execute(f'select * from {name_table_update}').description
            name_column = ()
            for i in list_name_column:
                name_column += (i[0],)
            # кортеж пустых значений для добавления в пустую строку
            list_values = ('',) * len(name_column)
            # вносим изменения
            cur.execute('''INSERT INTO {} {} VALUES {}'''.format(name_table_update, name_column, list_values))
            conn.commit()
        cur.close()


# удаляем строку из таблицы
def update_delete_row(list_db_update, name_table_update, number_row_delete):
    for name_db in list_db_update:
        # подключаемся к БД
        conn = sqlite3.connect(f'{os.path.abspath(os.getcwd())}\\DB\\{name_db}')
        cur = conn.cursor()
        # получаем ROWID по номеру строки (row_number_update)
        list_row_id = []
        for row_id in cur.execute('''SELECT ROWID FROM {}'''.format(name_table_update)).fetchall():
            list_row_id.append(row_id[0])
        # определяем  БД
        if cur.execute('''SELECT name FROM sqlite_master WHERE type='table' AND name="{}"'''.format(
                name_table_update)).fetchone():
            cur.execute(
                '''DELETE FROM {} WHERE ROWID="{}"'''.format(name_table_update, list_row_id[number_row_delete - 1]))
            conn.commit()
        cur.close()
