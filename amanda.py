import pandas as pd
import re
from datetime import datetime

# Настройка чтобы можно было увидеть все колонки
pd.set_option('display.max_columns', None)


def extract_main_data(text: str, key):
    """
    Функция для извлечения основных данных из колонки Main_data базового датафрейма.
    :param text: Текст разделенный символами переноса строки
    :return: Словарь
    """
    # Создаем список делая сплит по символу переноса строки
    lst_main_data = text.split('\n')
    # Создаем словарь с помощью генератора списков
    # Пришлось добавить условие по отбору строк, так как из за разметки бывает что получается новая строка без : и соответственно вылетает ошибка
    dict_main_data = {row.split(':')[0]: row.split(':')[1].strip() for row in lst_main_data if ':' in row}
    # print(dict_main_data.items())
    return dict_main_data.get(key)


def processing_capital(text):
    """
    Функция для обработки размера уставного капитала
    нужно учитывать следующие случаи : тыс,млн,млрд и отсутствие окончания,
    :param text:
    :return:
    """
    # Сначала сплитим по пробелу, затем конвертируем первый элемент во флоат, после чего умножаем на множитель исходя из
    # значения второго элемента
    if text:
        if text == '0':
            return 0
        else:
            digit, multiplier = text.split()
            clean_digit = float(digit)
            if 'тыс' in multiplier:
                clean_multipler = 1000
            elif 'млн' in multiplier:
                clean_multipler = 1000000
            elif 'млрд' in multiplier:
                clean_multipler = 1000000000
            else:
                clean_multipler = 1

            authorized_capital = clean_digit * clean_multipler
            return authorized_capital
    else:
        return 0


def mining_sait(text):
    # Воспользуемся регулярками.Если сайтов несколько то склеим их в строку через join
    lst_site = re.findall(r'[Сайт]{4}:\s(.+)', text)
    if lst_site == []:
        return ''
    else:
        return ','.join(lst_site)


def mining_okvad(text):
    """
    Функция для извлечения из текста значения ОКВЭД
    :param text: сырой текст
    :return: очищенный ОКВЭД
    """
    if text == 'Неопределенно':
        return text
    elif 'title' in text:
        result = text.split('>')[1].strip()
    else:
        result = text.split(':')[1].strip()
    return result


def extract_reg_msp(text: str):
    """
    Функция для извлечения сведений о том состоит ли организация в реестре МСП
    :param text:
    :return:
    """
    if text == 'Неопределенно':
        return None
    else:
        tmp_lst = text.split(':')
        date_reg = re.search(r'\d{2}.\d{2}.\d{4}', tmp_lst[1]).group()
        return date_reg


def extract_date_reg(text):
    """
    Функция для того чтобы извлечь дату регистрации в приемлимом виде.Поскольку rstrip лихо отрубает 0 у дат оканчивающихся на 0
    :param text:
    :return:
    """

    if text:
        date_reg = re.search(r'\d{2}[.]\d{2}[.]\d{4}', text)
        if date_reg:
            return date_reg.group()
        else:
            return None
    else:
        return None


# Загружаем таблицы с сырыми данными
# dtype для инн установлен строкой чтобы лидирующий ноль не убирался при создании датафрейма
base_df = pd.read_excel('resources/data_org.xlsx', dtype={'INN': str})

# Создаем итоговый датафрейм с заранее определенными колонками
df = pd.DataFrame(columns=['Организация', 'Юридическое наименование', 'ИНН', 'Руководитель', 'Уставной капитал',
                           'Численность персонала', 'Количество учредителей', 'Дата регистрации', 'Статус',
                           'ОКВЭД', 'Состоит в реестре МСП', 'Дата регистрации в реестр МСП','Специальные налоговые режимы', 'Индекс',
                           'Адрес',
                           'Координаты', 'Юридический адрес', 'Телефон', 'Факс', 'E-mail', 'Сайт',
                           'КПП', 'ОКПО', 'ОГРН', 'ОКФС', 'ОКОГУ', 'ОКОПФ', 'ОКТМО', 'ФСФР', 'ОКАТО',
                           'Предприятия рядом'])

# Удаляем строки  в которых вообще нет значений.
base_df.dropna(inplace=True,axis=0,how='all')

# Обработка данных из столбца Main_data содержащего в себе основную информацию
df['Организация'] = base_df['Name']
# Возникли небольшие трудности с передачей аргументов в функцию, но вспомнил зато как работать с кортежами
df['Юридическое наименование'] = base_df['Main_data'].apply(extract_main_data,
                                                            args=('Полное юридическое наименование',))

df['ИНН'] = base_df['INN']
df['Руководитель'] = base_df['Main_data'].apply(extract_main_data, args=('Руководитель',))
df['Руководитель'] = df['Руководитель'].apply(lambda x: ' '.join(x.split()[-3:]))
df['Уставной капитал'] = base_df['Main_data'].apply(extract_main_data, args=('Уставной капитал',))
# Обрабатываем колонку уставный капитал конвертируя ее в число
df['Уставной капитал'] = df['Уставной капитал'].apply(processing_capital)
# Обрабатываем колонку численость персонала
df['Численность персонала'] = base_df['Main_data'].apply(extract_main_data, args=('Численность персонала',))
df['Количество учредителей'] = base_df['Main_data'].apply(extract_main_data, args=('Количество учредителей',))
df['Дата регистрации'] = base_df['Main_data'].apply(extract_date_reg)
# df['Дата регистрации'] = df['Дата регистрации'].apply(lambda x: x.rstrip('_x000D_') if x else x)
df['Дата регистрации'] = pd.to_datetime(df['Дата регистрации'])
# df['Дата регистрации'] = df['Дата регистрации'].apply(lambda x: datetime.strptime(x, '%d.%m.%Y'))
df['Статус'] = base_df['Main_data'].apply(extract_main_data, args=('Статус',))

# Обработка данных из столбца Contacts
df['Индекс'] = base_df['Contacts'].apply(extract_main_data, args=('Индекс',))
df['Адрес'] = base_df['Contacts'].apply(extract_main_data, args=('Адрес',))
df['Координаты'] = base_df['Contacts'].apply(extract_main_data, args=('GPS координаты',))
df['Юридический адрес'] = base_df['Contacts'].apply(extract_main_data, args=('Юридический адрес',))
df['Телефон'] = base_df['Contacts'].apply(extract_main_data, args=('Телефон',))
df['Факс'] = base_df['Contacts'].apply(extract_main_data, args=('Факс',))
df['E-mail'] = base_df['Contacts'].apply(extract_main_data, args=('E-mail',))
df['Сайт'] = base_df['Contacts'].apply(mining_sait)

# Обрабатываем ОКВЭД
df['ОКВЭД'] = base_df['Okvad'].apply(mining_okvad)

# Обрабатываем данные из столбца  Reest, означающие есть ли организация в реестре малых или средних предприятий.
df['Состоит в реестре МСП'] = base_df['Reestr'].apply(lambda x: 'Нет' if x == 'Неопределенно' else 'Да')
df['Дата регистрации в реестр МСП'] = base_df['Reestr'].apply(extract_reg_msp)
df['Дата регистрации в реестр МСП'] = pd.to_datetime(df['Дата регистрации в реестр МСП'], format='%d.%m.%Y')
df['Специальные налоговые режимы'] = base_df['Main_data'].apply(extract_main_data,args=('Специальные налоговые режимы',))

# Обработка данных из столбца реквизиты
df['КПП'] = base_df['Rekvizit'].apply(extract_main_data,args=('КПП',))
df['ОКПО'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОКПО',))
df['ОГРН'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОГРН',))
df['ОКФС'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОКФС',))
df['ОКОГУ'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОКОГУ',))
df['ОКОПФ'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОКОПФ',))
df['ОКТМО'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОКТМО',))
df['ФСФР'] = base_df['Rekvizit'].apply(extract_main_data,args=('ФСФР',))
df['ОКАТО'] = base_df['Rekvizit'].apply(extract_main_data,args=('ОКАТО',))
df['Предприятия рядом'] = base_df['Rekvizit'].apply(extract_main_data,args=('Предприятия рядом',))
# Убираем словосочетание - Посмотреть все на карте чтобы оставить только названия рядом находящихся предприятий.
# Может быть попробую потом побаловатся кластерами и прочими вещами
df['Предприятия рядом'] = df['Предприятия рядом'].apply(lambda x:x.rstrip('- Посмотреть все на карте') if x else x)


print(df.dtypes)
# К слову надо все таки запомнить разницу между apply и applymap
df = df.applymap(lambda x: x.rstrip('_x000D_') if type(x) == str else x)

print(df.head())

df.to_excel('База данных организаций Бурятии.xlsx', index=False)
