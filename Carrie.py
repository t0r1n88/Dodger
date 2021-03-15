from docxtpl import DocxTemplate
import csv
from docxcompose.composer import Composer
from docx import Document as Document_compose
import os

# Считываем csv файл, не забывая что екселввский csv разделен на самомо деле не запятыми а точкой с запятой
reader = csv.DictReader(open('resources/mcpk_data.csv'), delimiter=';')
# Конвертируем объект reader в список словарей, где каждый словарь это строка с данными
data = list(reader)

# Создаем в цикле документы
for row in data:
    # Создаем объект шаблона
    doc = DocxTemplate('resources/template_mcpk.docx')
    # Получаем данные из объекта data
    context = {'nominative_fio': row['nominative_fio'], 'genitive_fio': row['genitive_fio'], 'программа': row['программа']}
    doc.render(context)
    doc.save(f'{row["nominative_fio"]} .docx')

