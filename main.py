from docxtpl import DocxTemplate
import csv

# Считываем csv файл, не забывая что екселввский csv разделен на самомо деле не запятыми а точкой с запятой
reader = csv.DictReader(open('resourses/1.csv'), delimiter=';')
# Конвертируем объект reader в список словарей
data = list(reader)

# Создаем в цикле документы
for row in data:
    doc = DocxTemplate('resources/template.docx')
    context = {'lastname': row['lastname'], 'firstname': row['firstname'], 'number': row['number']}
    doc.render(context)
    doc.save(f'{row["lastname"]} {row["firstname"]}.docx')


