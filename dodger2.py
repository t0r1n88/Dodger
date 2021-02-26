from docxtpl import DocxTemplate
import csv
from docxcompose.composer import Composer
from docx import Document as Document_compose
import os


# Считываем csv файл, не забывая что екселввский csv разделен на самомо деле не запятыми а точкой с запятой
reader = csv.DictReader(open('resources/data.csv'), delimiter=';')
# Конвертируем объект reader в список словарей
data = list(reader)

# Создаем в цикле документы
for row in data:
    doc = DocxTemplate('resources/template.docx')
    number = ''
    if len(row['number']) == 2:
        number = '000'+row['number']
    elif len(row['number']) == 3:
        number = '00'+row['number']
    elif len(row['number']) == 4:
        number = '0' + row['number']
    else:
        number = row['number']
    context = {'lastname': row['lastname'], 'firstname': row['firstname'], 'number': number,
               'profession': row['profession'], 'date_expiry': row['date_expiry'], 'date_issue': row['date_issue'],
               'qualification': row['qualification'],
               'category': row['category'], 'name_prep': row['name_prep'], 'name_dir': row['name_dir'],
               'hour': row['hour'],'base': row['base'], 'begin': row['begin'], 'end': row['end']}
    doc.render(context)
    doc.save(f'{row["lastname"]} {row["firstname"]}.docx')