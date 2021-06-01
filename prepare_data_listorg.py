import pandas  as pd

# Импортируем датафрейм

df = pd.read_excel('resources/data_list.xlsx')


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



