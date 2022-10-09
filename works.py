# Необходимо скачать библиотеку для работы скрипта:
# pip install pandas
#
import datetime
import shutil
from pathlib import Path
import pandas as pd


# Очистка папок для повторного запуска скрипта
pathpathes = ['documents', 'folder1', 'folder2', 'folder3', 'errors']
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
docs_dict = dict()
for code, doc, num, send in zip(codes, docs, numbers, sender):
    if doc != 'акт,счет':
        # Создание файла с одним документом форматирования
        path_str_2 = f'КА_{code}_{doc}_{num}'
        docs_dict[path_str_2] = {'code': code, 'doc': doc, 'num': num, 'send': send}
        Path('documents/' + path_str_2).touch()
    else:
        # Создания файлов с двумя документами форматирования
        # Дополнительно с проверкой на разделение не только по запятой, зачищяем возможные пробелы
        # Обработка неформатного вида документов
        stat1, stat2 = map(str.strip, doc.split(','))
        path_str_1 = f'КА_{code}_{stat1}_{num}'
        docs_dict[path_str_1] = {'code': code, 'doc': stat1, 'num': num, 'send': send}
        Path('documents/' + path_str_1).touch()

        path_str_2 = f'КА_{code}_{stat2}_{num}'
        docs_dict[path_str_2] = {'code': code, 'doc': stat2, 'num': num, 'send': send}
        Path('documents/' + path_str_2).touch()


# Start's 2nd ex
# Прописываем пути/названия будущих папок
dir_name = "documents/"
file_folder1 = 'folder1/'
file_folder2 = 'folder2/'
file_folder3 = 'folder3/'
file_folder4 = 'errors/'

# Создание папок для задания + папка с ошибочными файлами
Path(file_folder1).mkdir()
Path(file_folder2).mkdir()
Path(file_folder3).mkdir()
Path(file_folder4).mkdir()

# Выносим диапазоны дат
date_low, date_up = datetime.date(2020, 6, 20), datetime.date(2020, 7, 10)

# Цикл распределения по папкам
for i_file, i_prop in docs_dict.items():
    # Копирование файла в папку 1 по первому условию
    if i_prop['send'] == 1:
        shutil.copy(dir_name + i_file, file_folder1 + i_file)
    # Копирование файла в папку 2 по второму условию
    if i_prop['code'][-1] == i_prop['code'][-2]:
        shutil.copy(dir_name + i_file, file_folder2 + i_file)

    # Копирование файла в папку 3 по третьему условию + Отработка невалидных дат
    # Невозможно подсчитать через datetime изза 31 июня в файле
    try:
        file_date = datetime.date(*map(int, i_prop['num'].split('.')[::-1]))
    except ValueError:
        Path(dir_name + i_file).replace(file_folder4 + i_file)
    else:
        if date_low <= file_date <= date_up:
            shutil.copy(dir_name + i_file, file_folder3 + i_file)
