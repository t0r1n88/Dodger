"""
Скрипт для склонения ФИО по падежам

Получаем на вход файл эксель с фио
Делим на 3 части с помощью сплит
Каждую часть по отдельности склоняем по падежам
Склеиваем обратно
Сохраняем в виде дополнительных столбцов в документе.


Примечание

Предусмотреть обработку необычных случаев
1) Не русские фамилии
"""
from pymorphy2 import MorphAnalyzer
import pandas as pd


def create_cases(lastname, firstname, patronymic, lst_cases,morph):
    """
    Функция для склонения ФИО по падежам
    :param lastname: Фамилия
    :param firstname: Имя
    :param patronymic: Отчество
    :param lst_cases: Список падежей по которым нужно просклонять слово
    :param morph: Экземпляр класса morph
    :return:3 словаря в каждом из которых по 6 вариантов склоняемого слова
    """
    # Создаем словари где ключом будет падеж а значение слово в соответсвтующем падеже , хотя правильнее было бы использовать словари, где ключом был бы падеж
    # Создаем с помощью генераторов словарей
    dct_lastname = {case: '' for case in lst_cases}
    dct_firstname = {case: '' for case in lst_cases}
    dct_patronymic = {case: '' for case in lst_cases}

    # Анализируем каждое из слов
    lastname_parsed = morph.parse(lastname)[0]
    firstname_parsed = morph.parse(firstname)[0]
    patronymic_parsed = morph.parse(patronymic)[0]

    # Перебираем список падежей и при каждой итерации добавляем в словарь по соответствующему ключу слово в текущем падеже
    for case in lst_cases:
        dct_lastname[case] = lastname_parsed.inflect({case}).word
        dct_firstname[case] = firstname_parsed.inflect({case}).word
        dct_patronymic[case] = patronymic_parsed.inflect({case}).word

    return dct_lastname, dct_firstname, dct_patronymic


def create_case_fio(dct_last, dct_first, dct_patr, lst_cases):
    """
    Функция для объединения склоняемых слов в ФИО
    :param dct_last: Словарь с фамилиями
    :param dct_first: Словарь с именами
    :param dct_patr: Словарь с отчествами
    :param lst_cases: список падежей
    :return: словарь с просклонянеными ФИО
    """
    # Создаем словарь где ключ это падеж а значение ФИО в соответствующем падеже
    dct_fio = {case: '' for case in lst_cases}
    # Добавляем данные в словарь
    for case in lst_cases:
        dct_fio[case] = f'{dct_last[case]} {dct_first[case]} {dct_patr[case]}'.title()

    return dct_fio


base_df = pd.read_excel('resources/fio.xlsx')
# Создаем объект для морфологического анализа
morph = MorphAnalyzer()
test = []
# Список падежей
lst_cases = ['nomn', 'gent', 'datv', 'accs', 'ablt', 'loct']
for row in base_df.itertuples():
    # Создаем список из строки,делим по пробелам
    lastname, firstname, patronymic = row[1].split()
    gender = 'masc' if row[2] == 1 else 'femn'
    # Склоняем слова
    # dct_lastname, dct_firstname, dct_patronymic = create_cases(lastname, firstname, patronymic, lst_cases,morph)
    # value_to_table = create_case_fio(dct_lastname, dct_firstname, dct_patronymic, lst_cases)
    # test.append(value_to_table)

# # Создаем итоговый датафрейм
# df = pd.DataFrame(test)
# df.columns = ['Именительный', 'Родительный', 'Дательный', 'Винительный', 'Творительный', 'Предложный']
# # Сохраняем полученный датафрейм
# df.to_excel('fio_cases.xlsx',index=False)
