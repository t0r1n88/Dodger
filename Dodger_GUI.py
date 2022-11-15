import tkinter
import sys
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import string

from imap_tools import MailBox, AND
from xls2xlsx import XLS2XLSX
import os
from openpyxl import load_workbook
import pandas as pd
import tempfile
import time
# pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)



def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end
    path_to_end = filedialog.askdirectory()


def getMergedCellVal(sheet, cell):
    """
    Функция для получения значения объединеной ячейки
    Взято отсюда https://stackoverflow.com/questions/23562366/how-to-get-value-present-in-a-merged-cell
    """
    rng = [s for s in sheet.merged_cells.ranges if cell.coordinate in s]
    return sheet.cell(rng[0].min_row, rng[0].min_col).value if len(rng)!=0 else cell.value


def processing_data():
    """
    Фугкция для обработки данных
    :return:
    """
    not_used = ['Спам', 'Отправленные', 'Черновики', 'Корзина']
    cols_df = list(range(23))
    df = pd.DataFrame(columns=cols_df)

    with tempfile.TemporaryDirectory() as temp_dir:

        with MailBox('imap.mail.ru').login('myschool@copp03.ru', 'irjkf@_22') as mailbox:
            for f in mailbox.folder.list():
                if f.name not in not_used:
                    mailbox.folder.set(f.name)
                    for msg in mailbox.fetch():
                        # print(f' Subject {msg.subject}') # заголовок письма
                        # print(f' From {msg.from_}') # адрес почты отправителя
                        # print(f' Date {msg.date}') # время отправки
                        #
                        # print('***************')
                        msg_from = msg.from_  # получаем адрес почты отправителя
                        for att in msg.attachments:
                            if att.filename.endswith('.xlsx') or att.filename.endswith(
                                    '.xls'):  # проверяем на расширение

                                # work_file_name = att.filename.replace('.xls', '') # получаем название файла для варианта с xls
                                # print(work_file_name)
                                if att.filename.endswith('.xlsx'):  # Сохраняем во временную папку
                                    work_file_name = att.filename
                                    with open(f'{temp_dir}{att.filename}', 'wb') as f:
                                        f.write(att.payload)
                                elif att.filename.endswith('.xls'):  # конвертируем и сохраняем
                                    work_file_name = att.filename.replace('.xls', '.xlsx')
                                    with open(f'{temp_dir}{att.filename}', 'wb') as f:
                                        f.write(att.payload)
                                    out = XLS2XLSX(f'{temp_dir}{att.filename}')  # конвертируем в xlsx
                                    out.to_xlsx((f'{temp_dir}{work_file_name}'))  # сохраняем
                                    os.remove(f'{temp_dir}{att.filename}')  # удаляем файл xls чтобы не мешался

                                wb = load_workbook(f'{temp_dir}{work_file_name}')

                                # if att.filename.endswith('.xlsx'):
                                #     wb = load_workbook(f'{temp_dir}{att.filename}') # Загружаем созданный файл в режиме чтения
                                # else:
                                #     wb = load_workbook(f'{temp_dir}{file_name}.xlsx')
                                first_list = wb.sheetnames[0]  # получаем первый лист
                                standard_str = 'На обработку моих персональных данных в целях подключения к Личному кабинету в gosuslugi.ru:'  # проверочная строка

                                check_file = getMergedCellVal(wb[first_list], wb[first_list][
                                    'L2'])  # получаем значение ячейки,если совпадает то файл нужный нам
                                if check_file == standard_str:
                                    if len(wb.sheetnames) == 1:  # Проверяем длину
                                        name_org = wb[first_list]['B5'].value  # получаем значение ячейки B5
                                        print(name_org)
                                        temp_df = pd.read_excel(f'{temp_dir}{work_file_name}', skiprows=4,
                                                                header=None)  # считываем датафрейм
                                        temp_df[0] = msg_from
                                    else:
                                        len_sheets = len(wb.sheetnames)
                                        temp_df = pd.DataFrame(columns=list(range(23)))
                                        for sheet in wb.sheetnames:
                                            ml_temp_df = pd.read_excel(f'{temp_dir}{work_file_name}', sheet_name=sheet,
                                                                       skiprows=4, header=None)
                                            try:
                                                check_cols = ml_temp_df.iloc[:,
                                                             1].any()  # если есть хоть одно значение в колоноке 1 то добавляем эти данные
                                                if check_cols:
                                                    ml_temp_df[0] = msg_from
                                                    temp_df = pd.concat([temp_df, ml_temp_df], ignore_index=True)
                                            except IndexError:
                                                continue

                                    df = pd.concat([df, temp_df], ignore_index=True)

                                    if name_org:  # Сохраняем файл если есть имя организации
                                        name_org = name_org.translate(str.maketrans('', '',
                                                                                    string.punctuation))  # удаляем знаки препинания,которые могут помешать сохранить файлы
                                        wb.save(
                                            f'{path_to_end}/{name_org}.xlsx')  # Сохраняем файл под названием организации
                                    else:  # если не заполнено то сохраняем под емайлом откуда прислан файл.
                                        wb.save(f'{path_to_end}/{msg_from}.xlsx')

                            else:
                                continue

    t = time.localtime()
    current_time = time.strftime('%H_%M_%S', t)

    df.rename(columns={0: 'Откуда прислан файл', 1: 'Название учреждения'}, inplace=True)
    df.to_excel(f'{path_to_end}/Данные организаций для ФГИС Моя Школа от {current_time}.xlsx', index=False)
    messagebox.showinfo(message='Обработка завершена успешно!!!')



if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x560')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_report_6 = ttk.Frame(tab_control)
    tab_control.add(tab_report_6, text='Скрипт №1')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_report_6,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_report_6,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    btn_choose_end_folder = Button(tab_report_6, text='1) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder
                                       )
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    #Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_report_6, text='2) Получить данные', font=('Arial Bold', 20),
                                       command=processing_data
                                       )
    btn_proccessing_data.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()