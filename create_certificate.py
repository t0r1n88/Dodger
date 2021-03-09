from docxtpl import DocxTemplate
import csv



# Считываем csv файл, не забывая что екселввский csv разделен на самомо деле не запятыми а точкой с запятой
reader = csv.DictReader(open('resources/data_certificate.csv'), delimiter=';')
# Конвертируем объект reader в список словарей
data = list(reader)

# Создаем в цикле документы
for row in data:
    doc = DocxTemplate('resources/template_certificate.docx')

    context = {'dative_case_lastname': row['dative_case_lastname'], 'dative_case_firstname': row['dative_case_firstname'],
               'time': row['time'],
               'category_program': row['category_program'], 'format_program': row['format_program'],
               'name_program': row['name_program'],
               'hour': row['hour'],
               'chief_copp': row['chief_copp'], 'city': row['city'], 'year': row['year']
               }
    doc.render(context)
    doc.save(f'{row["dative_case_lastname"]} {row["dative_case_firstname"]}.docx')