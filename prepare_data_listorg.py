import pandas  as pd
import re
from styleframe import StyleFrame,Styler

"""
К слову вытаскивать все эти данные можно было и через регулярки
"""

def mining_okvad(text):
    """
    Функция для извлечения из текста значения ОКВЭД
    :param text: сырой текст
    :return: очищенный ОКВЭД
    """
    if 'title' in text:
        result = text.split('>')[1].strip()
    else:
        result = text.split(':')[1].strip()
    return result


def mining_email(text):
    """
    Функция для извлечения емайл и телефона организации. К слову наверное по правильному было бы делать функции по отдельности
    для каждого элемента, ведь из этого текста можно добывать координаты и полный адрес и прочие штуки
    :param lst_text: текст который мы потом сплитим по символу переноса
    :return:
    """

    lst_text = text.split('\n')
    for contact in lst_text:
        if 'E-mail' in contact:
            email = contact.split(':')[1].strip()
            return email
    # на случай если добрые верстальщики забудут прописать имэйл
    return ''


def mining_phone(text):
    lst_text = text.split('\n')

    for contact in lst_text:
        if 'Телефон' in contact:
            phone = contact.split(':')[1].strip()
            return phone
    # на случай если добрые верстальщики забудут прописать поле телефона
    return ''


def mining_sait(text):
    # Воспользуемся регулярками.Если сайтов несколько то склеим их в строку через join
    lst_site = re.findall(r'[Сайт]{4}:\s(.+)',text)
    if lst_site == []:
        return ''
    else:
        return ','.join(lst_site)


def mining_adress(text):
    # Функция для извлечения адреса
    lst_text = text.split('\n')

    for contact in lst_text:
        if 'Юридический адрес' in contact:
            ur_adress = contact.split(':')[1].strip()
            return ur_adress
    # на случай если добрые верстальщики забудут прописать поле телефона
    return ''



def processing_capital(text):
    """
    Функция для обработки размера уставного капитала
    нужно учитывать следующие случаи : тыс,млн,млрд и отсутствие окончания,
    :param text:
    :return:
    """
    # Сначала сплитим по пробелу, затем конвертируем первый элемент во флоат, после чего умножаем на множитель исходя из
    # значения второго элемента
    if text == 0:
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

        authorized_capital = clean_digit*clean_multipler
        return authorized_capital

def processing_category(size_human):
    """Функция для установки категории в зависимости от численности персонала
     5 категорий: микропредприятие, малое,среднее,большое,данные отсутствуют"""
    if size_human == 0:
        return 'Данные отсутствуют'
    elif size_human <=15:
        return 'Микропредприятие'
    elif size_human <=100:
        return 'Малое предприятие'
    elif size_human <=250:
        return 'Среднее предприятие'
    else:
        return 'Большое предприятие'




df = pd.read_excel('resources/data_list.xlsx')
out_df = pd.read_excel('resources/out_list_org.xlsx')



# Создаем новые столбцы
df['Телефон'] = df['Контакты'].apply(mining_phone)
df['E-mail'] = df['Контакты'].apply(mining_email)
df['ОКВЭД_чистый'] = df['ОКВЭД'].apply(mining_okvad)
df['Юридический адрес'] = df['Контакты'].apply(mining_adress)
df['Сайт'] = df['Контакты'].apply(mining_sait)

# Обрабатываем имеющиеся столбцы
# Обрабатываем уставной капитал, чтобы потом можно было с ним работать

df['Уставный капитал'] = df['Уставный капитал'].apply(processing_capital)

# Обрабатываем категории
df['Категория'] = df['Численость персонала'].apply(processing_category)



# Удаляем колонки
# Создаем базовый датафрейм
base_df = df.drop(['Название','ОКВЭД','Краткая справка','Контакты','УСН'],axis=1)

# Заполняем датафрейм для аналитиков
# Зполняем подпункты последовательностью, конечную цифру берем из данных датафрейма
# Не забываем что последовательность не включает в себя последнию цифру, поэтому и плюсуем 1
out_df['№ п/п'] = range(1,base_df.shape[0]+1)
out_df['Наименование предприятия/ организации'] = base_df['Юр.название']
out_df['ИНН'] = base_df['ИНН']
out_df['Основной вид экономической деятельности (по ОКВЭД 2)'] = base_df['ОКВЭД_чистый']
out_df['Среднесписочная численность работников'] = base_df['Численость персонала']
out_df['E-mail'] = base_df['E-mail']
out_df['Контакты (ФИО телефон) приемной, службы управления персоналом'] = base_df['Телефон']

# Используем библиотеку StyleFrame
excel_writer = StyleFrame.ExcelWriter('CAPL_table_for_analysis.xlsx')
sf = StyleFrame(out_df)

sf.apply_style_by_indexes(sf[(sf['Среднесписочная численность работников'] <=15) & (sf['Среднесписочная численность работников'] >=1)],cols_to_style='Среднесписочная численность работников',
                          styler_obj=Styler(bg_color='green'))

sf.apply_style_by_indexes(sf[sf['Среднесписочная численность работников'] ==0],cols_to_style='Среднесписочная численность работников',
                          styler_obj=Styler(bg_color='red'))

sf.apply_style_by_indexes(sf[(sf['Среднесписочная численность работников'] >15) & (sf['Среднесписочная численность работников'] <=100)],cols_to_style='Среднесписочная численность работников',
                          styler_obj=Styler(bg_color='orange'))
sf.apply_style_by_indexes(sf[(sf['Среднесписочная численность работников'] >100) & (sf['Среднесписочная численность работников'] )],cols_to_style='Среднесписочная численность работников',
                          styler_obj=Styler(bg_color='blue'))
#
#
#
sf.to_excel(excel_writer=excel_writer)
excel_writer.save()
# # out_df.to_excel('mintrud_df.xlsx',index=False)

