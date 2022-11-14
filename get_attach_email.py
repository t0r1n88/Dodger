"""
Получение вложений из почты для дальнейшей обработки

"""
from imap_tools import MailBox, AND

# Get date, subject and body len of all emails from INBOX folder
with MailBox('imap.mail.ru').login('myschool@copp03.ru', 'irjkf@_22') as mailbox:
# with MailBox('imap.mail.ru').login('budaev_oleg@copp03.ru', 'jsBcBmB9Jb1c6NHgR2eE') as mailbox:
    for msg in mailbox.fetch():
        for att in msg.attachments:
            if att.filename.endswith('.xlsx') or att.filename.endswith('.xls'):
                print(att.filename, att.content_type)
                with open('C:/3/{}'.format(att.filename), 'wb') as f:
                    f.write(att.payload)

    # for msg in mailbox.fetch():

        # for att in msg.attachments:
        #     if att.filename.endswith('.xlsx'):
        #         print(att.filename, att.content_type)
        #         with open('C:/3/{}'.format(att.filename), 'wb') as f:
        #             f.write(att.payload)