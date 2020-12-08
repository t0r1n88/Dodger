"""
в word документ вставить 
{{ director }}
далее будет подстановка из словаря context
"""

from docxtpl import DocxTemplate
import openpyxl

test=[]
wb = openpyxl.load_workbook('zaprosi-.xlsx')
sheet=wb.get_active_sheet()

for row in sheet['B3':'L5']:
    for cellObj in row:
        if cellObj.value==None or cellObj.value==" ":
            continue
        #print(cellObj.value)
        test.append(cellObj.value)
#print(test)

x=0
while x<len(test):
    doc = DocxTemplate("шаблон-запрос.docx")
    context = { 'emitent' : test[x],'address1' : test[x+5],'участник' : test[x+1],'адрес_участника' : test[x+4],'dolznost' : test[x+10]
                ,'director' : test[x+9]}
    doc.render(context)
    doc.save(str(test[x].replace("\"",""))+"-"+str(test[x+9])+'.doc')    
    x+=11

