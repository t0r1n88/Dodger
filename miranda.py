from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk
from pymorphy2 import MorphAnalyzer
import pandas as pd


def select_file_template_contracts():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_contracts
    name_file_template_contracts = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_contracts():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_contracts
    name_file_data_contracts = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_contracts():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_contracts
    path_to_end_folder_contracts = filedialog.askdirectory()


def create_case(file_fio):
    """
    Функция для склонения фио в родительном падеже
    :param data:файл с данными
    :return:словарь вида: nomn:ФИО в именительном падеже, gent: ФИО в родительном падеже
    """
    # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
    reader = csv.DictReader(open(file_fio), delimiter=';')
    # Конвертируем объект reader в список словарей
    data = list(reader)
    for row in data:
        lastname, firstname, patronymic = row['ФИО_слушателя'].split()
        print(lastname, firstname, patronymic)


def generate_contracts():
    """
    Функция для создания договоров
    :return:
    """
    try:
        # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
        reader = csv.DictReader(open(name_file_data_contracts), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)
        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_contracts)
            context = {'ФиоСлушателя': row['ФиоСлушателя'], 'НомерДоговора': row['НомерДоговора'],
                       'ДатаПодписанияДоговора': row['ДатаПодписанияДоговора'],
                       'ДолжностьФиоРодительныйПадеж': row['ДолжностьФиоРодительныйПадеж'],
                       'программа': row['программа'], 'срок_в_месяцах': row['срок_в_месяцах'],
                       'профессия': row['профессия'], 'срок_в_часах': row['срок_в_часах'],
                       'дата_начала_занятий': row['дата_начала_занятий'],
                       'начало_обучения': row['начало_обучения'],
                       'конец_обучения': row['конец_обучения'], 'полная_стоимость': row['полная_стоимость'],
                       'первая_часть_оплаты': row['первая_часть_оплаты'],
                       'дата_первой_оплаты': row['дата_первой_оплаты'],
                       'вторая_часть_оплаты': row['вторая_часть_оплаты'],
                       'дата_второй_оплаты': row['дата_второй_оплаты'],
                       'третья_часть_оплаты': row['третья_часть_оплаты'],
                       'дата_третьей_оплаты': row['дата_третьей_оплаты'],
                       'дата_окончания_договора': row['дата_окончания_договора'],
                       'должность_подписывающего': row['должность_подписывающего'],
                       'ФИО_подписывающего': row['ФИО_подписывающего'],
                       'дата_подписи_договора': row['дата_подписи_договора'], 'дата_рождения': row['дата_рождения'],
                       'серия_паспорта': row['серия_паспорта'], 'номер_паспорта': row['номер_паспорта'],
                       'дата_выдачи_паспорта': row['дата_выдачи_паспорта'], 'выдан': row['выдан'],
                       'адрес_регистрации': row['адрес_регистрации'], 'снилс': row['снилс'],
                       'контактный_телефон': row['контактный_телефон']}
            doc.render(context)
            doc.save(f'{path_to_end_folder_contracts}/{row["ФИОслушателя"]}.docx')
        messagebox.showinfo('Dodger', 'Создание свидетельств успешно завершено!')
    except NameError as e:
        print(e)
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


# Создаем окно
if __name__ == '__main__':
    window = Tk()
    window.title('Dodger')
    window.geometry('640x480')

    # Создаем ФИО в родительском падеже
    # dct_genitive_fio = create_case(name_file_data_contract)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку свидетельства о повышении
    tab_contract = ttk.Frame(tab_control)
    tab_control.add(tab_contract, text='Создание договоров')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_contract, text='Скрипт для создания договоров')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон

    btn_template_contract = Button(tab_contract, text='1) Выберите шаблон договора', font=('Arial Bold', 20),
                                   command=select_file_template_contracts
                                   )
    btn_template_contract.grid(column=0, row=1, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_contract = Button(tab_contract, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                               command=select_file_data_contracts
                               )
    btn_data_contract.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_contract = Button(tab_contract, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                            command=select_end_folder_contracts
                                            )
    btn_choose_end_folder_contract.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для запуска функции генерации файлов

    btn_create_files_contract = Button(tab_contract, text='4) Создать договора', font=('Arial Bold', 20),
                                       command=generate_contracts
                                       )
    btn_create_files_contract.grid(column=0, row=4, padx=10, pady=10)

    #
    #
    # # Создаем вкладку для создания  сертификатов
    # tab_scc = ttk.Frame(tab_control)
    # tab_control.add(tab_scc, text='Создание сертификатов')
    #
    # # Добавляем виджеты на вкладку
    # lbl_hello = Label(tab_scc, text='Скрипт для создания сертификатов')
    # lbl_hello.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Создаем кнопку Выбрать шаблон
    #
    # btn_template_scc = Button(tab_scc, text='1) Выберите шаблон сертификата', font=('Arial Bold', 20),
    #                           command=select_file_template_scc, )
    # btn_template_scc.grid(column=0, row=1, padx=10, pady=10)
    #
    # # Создаем кнопку Выбрать файл с данными
    # btn_data_scc = Button(tab_scc, text='2) Выберите файл с данными', font=('Arial Bold', 20),
    #                       command=select_file_data_scc)
    # btn_data_scc.grid(column=0, row=2, padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_folder_scc = Button(tab_scc, text='3) Выберите конечную папку', font=('Arial Bold', 20),
    #                                    command=select_end_folder_scc)
    # btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Создаем кнопку для запуска функции генерации файлов
    #
    # btn_create_files_scc = Button(tab_scc, text=' Создать сертификаты', font=('Arial Bold', 20),
    #                               command=generate_scc)
    # btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)
    #
    #
    #
    #
    #
    #
    # # Создаем вкладку Создание удостоверений
    # tab_certificate = ttk.Frame(tab_control)
    # tab_control.add(tab_certificate, text='Создание удостоверений')
    #
    # # Добавляем виджеты на вкладку
    # lbl_hello = Label(tab_certificate, text='Скрипт для создания удостоверений')
    # lbl_hello.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Создаем кнопку Выбрать шаблон
    #
    # btn_template_scc = Button(tab_certificate, text='1) Выберите шаблон удостоверения', font=('Arial Bold', 20),
    #                           command=select_file_template_certificates )
    # btn_template_scc.grid(column=0, row=1, padx=10, pady=10)
    #
    # # Создаем кнопку Выбрать файл с данными
    # btn_data_scc = Button(tab_certificate, text='2) Выберите файл с данными', font=('Arial Bold', 20),
    #                       command=select_file_data_certificates)
    # btn_data_scc.grid(column=0, row=2, padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_folder_scc = Button(tab_certificate, text='3) Выберите конечную папку', font=('Arial Bold', 20),
    #                                    command=select_end_folder_certificates)
    # btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Создаем кнопку для запуска функции генерации файлов
    #
    # btn_create_files_scc = Button(tab_certificate, text=' Создать удостоверения', font=('Arial Bold', 20),
    #                               command=generate_certificates)
    # btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()
