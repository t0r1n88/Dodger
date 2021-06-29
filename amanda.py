import pandas as pd

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
# Загружаем таблицы с сырыми данными
# dtype для инн установлен строкой чтобы лидирующий ноль не убирался при создании датафрейма
base_df = pd.read_excel('resources/data_org.xlsx', dtype={'INN': str})

# Создаем итоговый датафрейм с заранее определенными колонками
df = pd.DataFrame(columns=['Организация', 'Юридическое наименование', 'ИНН', 'Руководитель', 'Уставной капитал',
                           'Численность персонал', 'Количество учредителей', 'Дата регистрации', 'Статус'])

# Записываем данные из столбцов Name и ИНН в итогоый датафрейм
df['Организация'] = base_df['Name']
# Возникли небольшие трудности с передачей аргументов в функцию, но вспомнил зато как работать с кортежами
df['Юридическое наименование'] = base_df['Main_data'].apply(extract_main_data,
                                                            args=('Полное юридическое наименование',))
df['ИНН'] = base_df['INN']
df['Руководитель'] = base_df['Main_data'].apply(extract_main_data, args=('Руководитель',))
# Теперь нужно очистить столбец Руководитель от наименование должности оставив только ФИС
# df['Руководитель'] = df['Руководитель'].apply(lambda x:x.upper().split('ДИРЕКТОР') if 'ДИРЕКТОР' in x.upper())
# Хотя можно сделать все намного проще и просто оставлять последие 3 слова
df['Руководитель'] = df['Руководитель'].apply(lambda x: ' '.join(x.split()[-3:]))
df['Уставной капитал'] = base_df['Main_data'].apply(extract_main_data, args=('Уставной капитал',))
# Обрабатываем колонку уставный капитал конвертируя ее в число
df['Уставной капитал'] = df['Уставной капитал'].apply(processing_capital)
print(df[['Организация', 'Юридическое наименование', 'ИНН', 'Руководитель', 'Уставной капитал']].head())
# df.to_excel('База данных организаций Бурятии.xlsx',index=False)
