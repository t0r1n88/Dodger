"""
в word документ вставить 
{{ director }}
далее будет подстановка из словаря context
"""

from docxtpl import DocxTemplate
import openpyxl
import os

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
    doc = DocxTemplate("шаблон-ответ.docx")
    context = {'участник' : test[x+1],'адрес_участника' : test[x+4], 'общество' : test[x],'адрес_общества' : test[x+5],'dolznost' : test[x+3]
                ,'director' : test[x+2]}
    doc.render(context)    
    doc.save(test[x+1].replace("\"","")+"-"+test[x+2]+"-"+test[x].replace("\"","")+'.doc')    
    x+=11


