"""
Получение вложений из почты для дальнейшей обработки

"""
from imap_tools import MailBox, AND
from xls2xlsx import XLS2XLSX
import os


path_to_end = 'C:/Данные/'

# Get date, subject and body len of all emails from INBOX folder
not_used = ['Спам','Отправленные','Черновики','Корзина']


with MailBox('imap.mail.ru').login('myschool@copp03.ru', 'irjkf@_22') as mailbox:
    for f in mailbox.folder.list():
        if f.name not in not_used:
            print(f.name)
            mailbox.folder.set(f.name)
            for msg in mailbox.fetch():
                print(f' Subject {msg.subject}') # заголовок письма
                print(f' From {msg.from_}') # адрес почты отправителя
                print(f' Date {msg.date}') # время отправки

                print('***************')
                for att in msg.attachments:
                    if att.filename.endswith('.xlsx'):  # 2007 excel
                        # print(att.filename, att.content_type)

                        with open(f'{path_to_end}{att.filename}','wb') as f:
                            f.write(att.payload)
                    elif att.filename.endswith(('.xls')):  # 2003 excel
                        file_name = att.filename.replace('.xls', '')
                        with open(f'{path_to_end}{att.filename}', 'wb') as f:
                            f.write(att.payload)
                        out = XLS2XLSX(f'{path_to_end}{att.filename}') # конвертируем в
                        out.to_xlsx((f'{path_to_end}{file_name}.xlsx')) # сохраняем
                        os.remove(f'{path_to_end}{att.filename}') # удаляем файл xls чтобы не мешался
                    else:
                        continue
#
#
# # with MailBox('imap.mail.ru').login('budaev_oleg@copp03.ru', 'jsBcBmB9Jb1c6NHgR2eE') as mailbox:
#     for msg in mailbox.fetch():
#         print(f' Subject {msg.subject}') # заголовок письма
#         print(f' From {msg.from_}') # адрес почты отправителя
#         print(f' Date {msg.date}') # время отправки
#
#         print('***************')
#         for att in msg.attachments:
#             if att.filename.endswith('.xlsx'):  # 2007 excel
#                 # print(att.filename, att.content_type)
#
#                 with open(f'{path_to_end}{att.filename}','wb') as f:
#                     f.write(att.payload)
#             elif att.filename.endswith(('.xls')):  # 2003 excel
#                 file_name = att.filename.replace('.xls', '')
#                 with open(f'{path_to_end}{att.filename}', 'wb') as f:
#                     f.write(att.payload)
#                 out = XLS2XLSX(f'{path_to_end}{att.filename}') # конвертируем в
#                 out.to_xlsx((f'{path_to_end}{file_name}.xlsx')) # сохраняем
#                 os.remove(f'{path_to_end}{att.filename}') # удаляем файл xls чтобы не мешался
#             else:
#                 continue

    # for msg in mailbox.fetch():

