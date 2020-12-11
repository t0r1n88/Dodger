from docx import Document

# Создаем пустой вордовский документ
document = Document()
"""
Так как все в ворде это параграфы добавляем параграф для примера
"""
# Присваиваем ссылку чтобы можно было добавлять параграфы сверху или снизу
name_copp = document.add_paragraph('Центр опережающей профессиональной подготовки Республики Бурятия')
# Вставляем параграф сверху
above_paragraph = name_copp.insert_paragraph_before('Во славу Знания!!!')

document.save('Prototype.docx')
