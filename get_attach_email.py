"""
Получение вложений из почты для дальнейшей обработки

"""
import string

from imap_tools import MailBox, AND
from xls2xlsx import XLS2XLSX
import os
from openpyxl import load_workbook
import pandas as pd

def getMergedCellVal(sheet, cell):
    """
    Функция для получения значения объединеной ячейки
    Взято отсюда https://stackoverflow.com/questions/23562366/how-to-get-value-present-in-a-merged-cell
    """
    rng = [s for s in sheet.merged_cells.ranges if cell.coordinate in s]
    return sheet.cell(rng[0].min_row, rng[0].min_col).value if len(rng)!=0 else cell.value



temp_dir = 'C:/temp/'
path_to_end = 'C:/Данные/'
# Get date, subject and body len of all emails from INBOX folder
not_used = ['Спам','Отправленные','Черновики','Корзина']


with MailBox('imap.mail.ru').login('myschool@copp03.ru', 'irjkf@_22') as mailbox:
    for f in mailbox.folder.list():
        if f.name not in not_used:
            # print(f.name)
            mailbox.folder.set(f.name)
            for msg in mailbox.fetch():
                # print(f' Subject {msg.subject}') # заголовок письма
                # print(f' From {msg.from_}') # адрес почты отправителя
                # print(f' Date {msg.date}') # время отправки
                #
                # print('***************')
                msg_from = msg.from_ # получаем адрес почты отправителя
                for att in msg.attachments:
                    if att.filename.endswith('.xlsx'):  # 2007 excel
                        with open(f'{temp_dir}{att.filename}','wb') as f:
                            f.write(att.payload)
                        wb = load_workbook(f'{temp_dir}{att.filename}') # Загружаем созданный файл в режиме чтения
                        first_list = wb.sheetnames[0] # получаем первый лист
                        standard_str = 'На обработку моих персональных данных в целях подключения к Личному кабинету в gosuslugi.ru:' # проверочная строка

                        check_file = getMergedCellVal(wb[first_list], wb[first_list]['L2']) # получаем значение ячейки,если совпадает то файл нужный нам
                        if check_file == standard_str:
                            name_org = wb[first_list]['B5'].value # получаем значение ячейки B5

                            print(name_org)
                            temp_df = pd.read_excel(f'{temp_dir}{att.filename}',skiprows=4,header=None) # считываем датафрейм

                            if name_org: # Сохраняем файл если есть имя организации
                                name_org = name_org.translate(str.maketrans('','',string.punctuation)) # удаляем знаки препинания,которые могут помешать сохранить файлы
                                wb.save(f'{path_to_end}{name_org}.xlsx') # Сохраняем файл под названием организации
                            else:
                                wb.save(f'{path_to_end}{msg_from}.xlsx')
                        #
                        #
                        # print(temp_df)


                    elif att.filename.endswith(('.xls')):  # 2003 excel
                        file_name = att.filename.replace('.xls', '')
                        with open(f'{temp_dir}{att.filename}', 'wb') as f:
                            f.write(att.payload)
                        out = XLS2XLSX(f'{temp_dir}{att.filename}') # конвертируем в
                        out.to_xlsx((f'{temp_dir}{file_name}.xlsx')) # сохраняем
                        os.remove(f'{temp_dir}{att.filename}') # удаляем файл xls чтобы не мешался
                    else:
                        continue
#


