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


def create_cases(lastname, firstname, patronymic, gender, lst_cases, morph=MorphAnalyzer()):
    """
    Функция для склонения ФИО по падежам
    :param lastname: Фамилия
    :param firstname: Имя
    :param patronymic: Отчество
    :param gender: Пол
    :param lst_cases: Список падежей по которым нужно просклонять слово
    :param morph: Экземпляр класса morph
    :return:3 словаря в каждом из которых по 6 вариантов склоняемого слова
    """
    # Создаем теги наличие которых будет говорить о возможности просклонять слово
    tag_lastname = 'Surn'
    tag_firstname = 'Name'
    tag_patr = 'Patr'
    # Создаем словари где ключом будет падеж а значение слово в соответсвтующем падеже , хотя правильнее было бы использовать словари, где ключом был бы падеж
    # Создаем с помощью генераторов словарей
    dct_lastname = {case: '' for case in lst_cases}
    dct_firstname = {case: '' for case in lst_cases}
    dct_patronymic = {case: '' for case in lst_cases}

    # Анализируем каждое из слов
    # lastname_parsed = parse_case(lastname)
    # firstname_parsed = parse_case(firstname)
    # patronymic_parsed = parse_case(patronymic)
    #
    # lastname_parsed = morph.parse(lastname)[0]
    # firstname_parsed = morph.parse(firstname)[0]
    # patronymic_parsed = morph.parse(patronymic)[0]

    # Перебираем список падежей и при каждой итерации добавляем в словарь по соответствующему ключу слово в текущем падеже
    for case in lst_cases:
        # dct_lastname[case] = lastname_parsed.inflect({case}).word
        # dct_firstname[case] = firstname_parsed.inflect({case}).word
        # dct_patronymic[case] = patronymic_parsed.inflect({case}).word
        dct_lastname[case] = parse_case(lastname, gender, case, tag_lastname, morph)
        dct_firstname[case] = parse_case(firstname, gender, case, tag_firstname, morph)
        dct_patronymic[case] = parse_case(patronymic, gender, case, tag_patr, morph)

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


def parse_case(word, gender, case, tag_fio, morph=MorphAnalyzer()):
    """
    Функция для проверки возможности склонения слова по роду и падежу
    :param word: проверяемое слово
    :param gender: проверяемый пол
    :param case: проверяемый падеж
    :param morph: анализатор
    :return: Возвращает слово в неизменненом виде если нельзя просклонять по падежу, если можно то измененное
    """
    # Парсим слово получаем все возможные лексемы
    word_parsed = morph.parse(word)
    # Перебираем полученные разборы на предмет совпадений
    for par in word_parsed:
        if tag_fio in par.tag:
            if par.inflect({gender,tag_fio,'sing'}):
                print(par.inflect({gender,case}).word)
                return par.inflect({case}).word
            return par.word
    return par.word


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
    dct_lastname, dct_firstname, dct_patronymic = create_cases(lastname, firstname, patronymic, gender, lst_cases,
                                                               morph)
    value_to_table = create_case_fio(dct_lastname, dct_firstname, dct_patronymic, lst_cases)
    test.append(value_to_table)

# Создаем итоговый датафрейм
df = pd.DataFrame(test)
df.columns = ['Именительный', 'Родительный', 'Дательный', 'Винительный', 'Творительный', 'Предложный']
# Сохраняем полученный датафрейм
df.to_excel('fio_cases.xlsx',index=False)
