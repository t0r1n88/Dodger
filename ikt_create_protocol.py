from docxtpl import DocxTemplate
import pandas as pd

"""
На будущее не забыть что максимальный балл может измениться и название этой колонки надо менять в зависимости от теста
(максимальный балл может менятся и соответсвенно название колонки будет меняться)
"""

# Считываем датафрейм
df = pd.read_excel('resources/ikt_temp.xlsx', usecols='A:K')
# Получаем максимальный балл из заголовка
# Оценка/25,00
trash_max_ball = df.columns[8].split('/')[1]
# конвертируем
max_ball = float(trash_max_ball.replace(',', '.'))
# Обработаем датафрейм  приведем к флоату столбец Оценка
df['Оценка/25,00'] = df['Оценка/25,00'].str.replace(',', '.')
df['Оценка/25,00'] = df['Оценка/25,00'].astype(float)
# Итерируемся по датафрейму с помощью itertupples, понятнее получается
for value in df.itertuples():
    doc = DocxTemplate('resources/Шаблон протокола на знание основ ИКТ.docx')
    fio = f'{value[1]} {value[2]}'

    context = {'ФИО': fio, 'Должность': value[4], 'НабБалл': value[9], 'Процент': (value[9] / max_ball) * 100,
               'МаксБалл': max_ball, 'ПрБалл': value[10], 'ПрПроцент': value[11], 'Длительность': value[8],
               'ДатаПроведения': value[7],
               'Результат': 'Тест Сдан' if value[9] > value[10] else 'Тест не сдан'}

    doc.render(context)
    doc.save(f'{context["ФИО"]}.docx')
