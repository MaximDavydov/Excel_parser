# Нужно скачать библиотеку для работы скрипта:
# pip install pandas
#
import os
import datetime
import shutil
from pathlib import Path
import pandas as pd


# Очистка папок для повторного запуска скрипта
pathpathes = ['documents', 'folder1',
              'folder2', 'folder3',
              'errors']
for i_path in pathpathes:
    try:
        shutil.rmtree(i_path)
    except FileNotFoundError:
        pass

# Start's 1st ex
Path("documents").mkdir()

list1 = pd.read_excel('Задание.xlsx')

codes = list1['Код_КА'].tolist()
docs = list1['Документы для формирования'].tolist()
numbers = list1['Дата документа'].tolist()
sender = list1['Отправить'].tolist()

# Расширенный список по данным
docs_ = []
for code, doc, num, send in zip(codes, docs, numbers, sender):
    if doc != 'акт,счет':
        # Создание файла с одним документом форматирования
        docs_list.append([code, doc, num, send])
        path_str_2 = f'documents/КА_{code}_{doc}_{num}'
        Path(path_str_2).touch()
    else:
        # Создания файлов с двумя документами форматирования
        # Дополнительно с проверкой на разделение не только по запятой, зачищяем возможные пробелы
        stat1, stat2 = map(str.strip, doc.split(','))
        docs_list.append([code, stat1, num, send])
        path_str_1 = f'documents/КА_{code}_{stat1}_{num}'
        Path(path_str_1).touch()

        docs_list.append([code, stat2, num, send])
        path_str_2 = f'documents/КА_{code}_{stat2}_{num}'
        Path(path_str_2).touch()

print(docs_list)
# print(codes)
# print(docs)
# print(numbers)
# print(sender)

# Start's 2nd ex
# Создание папок для задания + папка с ошибочными файлами
Path("folder1").mkdir()
Path("folder2").mkdir()
Path("folder3").mkdir()
Path("errors").mkdir()

dir_name = "documents/"
file_folder1 = 'folder1/'
file_folder2 = 'folder2/'
file_folder3 = 'folder3/'
file_folder4 = 'errors/'

# Читаем файлы из папки
get_files = os.listdir(dir_name)

# Цикл распределения по папкам
for i_item in docs_list:
    # Копирование файла в папку 1 по первому условию
    if i_item[-1] == 1:
        # g = get_files[docs_list.index(i_item)]
        g = f'КА_{i_item[0]}_{i_item[1]}_{i_item[2]}'
        # os.replace(dir_name + g, file_folder1 + g)
        shutil.copy(dir_name + g, file_folder1 + g)
    # Копирование файла в папку 2 по первому условию
    if i_item[0][-1] == i_item[0][-2]:
        g = f'КА_{i_item[0]}_{i_item[1]}_{i_item[2]}'
        shutil.copy(dir_name + g, file_folder2 + g)

    # Список с датой для сравнения диапазона даты
    file_date_list = list(map(int, i_item[2].split('.')[::-1]))

    # Невозможно подсчитать через datetime изза 31 июня в файле
    # Ненормализированные данные, mishit date - move to the other folder
    try:
        file_date = datetime.date(file_date_list[0], file_date_list[1], file_date_list[2])
    except ValueError:
        file_name = f'КА_{i_item[0]}_{i_item[1]}_{i_item[2]}'
        os.replace(dir_name + file_name, file_folder4 + file_name)
    else:
        if datetime.date(2020, 6, 20) <= file_date <= datetime.date(2020, 7, 10):
            file_name = f'КА_{i_item[0]}_{i_item[1]}_{i_item[2]}'
            shutil.copy(dir_name + file_name, file_folder3 + file_name)