from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk


# Функции выбора шаблона,файла с данными и конеченой папки для генерации свидетельств

def select_file_template_diplomas():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_diplomas
    name_file_template_diplomas = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_diplomas():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_diplomas
    name_file_data_diplomas = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_diplomas():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_diplomas
    path_to_end_folder_diplomas = filedialog.askdirectory()


# Функции выбора шаблона,файла с данными и конеченой папки для генерации сертификатов

def select_file_template_scc():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_scc
    name_file_template_scc = filedialog.askopenfilename(filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_scc():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_scc
    name_file_data_scc = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_scc():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_scc
    path_to_end_folder_scc = filedialog.askdirectory()


# Функции выбора шаблона,файла с данными и конеченой папки для генерации удостоверений

def select_file_template_certificates():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_certificates
    name_file_template_certificates = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_certificates():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_certificates
    name_file_data_certificates = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_certificates():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_certificates
    path_to_end_folder_certificates = filedialog.askdirectory()


def generate_diplomas():
    """
    Функция для создания свидетельств из шаблона и файла с данными

    """
    try:
        # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
        reader = csv.DictReader(open(name_file_data_diplomas), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)
        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_diplomas)
            # Код для того чтобы операторы вводили номера без нулей при заполнении таблицы с данными
            number = ''
            if len(row['НомерДок']) == 2:
                number = '000' + row['НомерДок']
            elif len(row['НомерДок']) == 3:
                number = '00' + row['НомерДок']
            elif len(row['НомерДок']) == 4:
                number = '0' + row['НомерДок']
            else:
                number = row['НомерДок']
            context = {'ФИОСлушателя': row['ФИОСлушателя'], 'Фамилия': row['Фамилия'], 'НомерДок': number,
                       'ИмяОтчество': row['ИмяОтчество'], 'ФИОСлушателяРодПадеж': row['ФИОСлушателяРодПадеж'],
                       'region_genitive_case': row['region_genitive_case'],
                       'Оператор': row['Оператор'], 'РегионРодПадеж': row['РегионРодПадеж'],
                       'ВидПрограммы': row['ВидПрограммы'], 'ПодвидПрограммы': row['ПодвидПрограммы'],
                       'ФорматПрограммы': row['ФорматПрограммы'], 'НазваниеПрограммы': row['НазваниеПрограммы'],
                       'Профессия': row['Профессия'], 'Квалификация': row['Квалификация'],
                       'КатегорияРазряд': row['КатегорияРазряд'],
                       'Группа': row['Группа'], 'ФИОПреп': row['ФИОПреп'], 'КолЧасовОбуч': row['КолЧасовОбуч'],
                       'КолМесяцевОбуч': row['КолМесяцевОбуч'], 'БазаОбучения': row['БазаОбучения'],
                       'ТекстНачалоОбуч': row['ТекстНачалоОбуч'],
                       'ТекстОконОбуч': row['ТекстОконОбуч'], 'ЧислоНачалоОбуч': row['ЧислоНачалоОбуч'],
                       'ЧислоКонецОбуч': row['ЧислоКонецОбуч'],
                       'ДолжностьФиоРодПадеж': row['ДолжностьФиоРодПадеж'],
                       'ПолнаяСтоимость': row['ПолнаяСтоимость'], 'ПерваяЧастьОплаты': row['ПерваяЧастьОплаты'],
                       'ДатаПервойОплаты': row['ДатаПервойОплаты'],
                       'ВтораяЧастьОплаты': row['ВтораяЧастьОплаты'], 'ДатаВторойОплаты': row['ДатаВторойОплаты'],
                       'ТретьяЧастьОплаты': row['ТретьяЧастьОплаты'],
                       'ДатаТретьейОплаты': row['ДатаТретьейОплаты'],
                       'ДатаОкончанияДоговора': row['ДатаОкончанияДоговора'],
                       'ДолжностьПодписывающего': row['ДолжностьПодписывающего'],
                       'ФиоПодписывающего': row['ФиоПодписывающего'],
                       'НазваниеОрганизации': row['НазваниеОрганизации'],
                       'Исполнитель': row['Исполнитель'], 'ДатаПодписанияДок': row['ДатаПодписанияДок'],
                       'Город': row['Город'],
                       'Год': row['Год']}
            doc.render(context)
            doc.save(f'{path_to_end_folder_diplomas}/{row["Фио"]}.docx')
        messagebox.showinfo('Dodger', 'Создание свидетельств успешно завершено!')
    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


def generate_scc():
    """
    Функция для создания сертификатов

    """
    try:
        reader = csv.DictReader(open(name_file_data_scc), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)

        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_scc)

            context = {'ФИОСлушателя': row['ФИОСлушателя'], 'Фамилия': row['Фамилия'], 'НомерДок': number,
                       'ИмяОтчество': row['ИмяОтчество'], 'ФИОСлушателяРодПадеж': row['ФИОСлушателяРодПадеж'],
                       'region_genitive_case': row['region_genitive_case'],
                       'Оператор': row['Оператор'], 'РегионРодПадеж': row['РегионРодПадеж'],
                       'ВидПрограммы': row['ВидПрограммы'], 'ПодвидПрограммы': row['ПодвидПрограммы'],
                       'ФорматПрограммы': row['ФорматПрограммы'], 'НазваниеПрограммы': row['НазваниеПрограммы'],
                       'Профессия': row['Профессия'], 'Квалификация': row['Квалификация'],
                       'КатегорияРазряд': row['КатегорияРазряд'],
                       'Группа': row['Группа'], 'ФИОПреп': row['ФИОПреп'], 'КолЧасовОбуч': row['КолЧасовОбуч'],
                       'КолМесяцевОбуч': row['КолМесяцевОбуч'], 'БазаОбучения': row['БазаОбучения'],
                       'ТекстНачалоОбуч': row['ТекстНачалоОбуч'],
                       'ТекстОконОбуч': row['ТекстОконОбуч'], 'ЧислоНачалоОбуч': row['ЧислоНачалоОбуч'],
                       'ЧислоКонецОбуч': row['ЧислоКонецОбуч'],
                       'ДолжностьФиоРодПадеж': row['ДолжностьФиоРодПадеж'],
                       'ПолнаяСтоимость': row['ПолнаяСтоимость'], 'ПерваяЧастьОплаты': row['ПерваяЧастьОплаты'],
                       'ДатаПервойОплаты': row['ДатаПервойОплаты'],
                       'ВтораяЧастьОплаты': row['ВтораяЧастьОплаты'], 'ДатаВторойОплаты': row['ДатаВторойОплаты'],
                       'ТретьяЧастьОплаты': row['ТретьяЧастьОплаты'],
                       'ДатаТретьейОплаты': row['ДатаТретьейОплаты'],
                       'ДатаОкончанияДоговора': row['ДатаОкончанияДоговора'],
                       'ДолжностьПодписывающего': row['ДолжностьПодписывающего'],
                       'ФиоПодписывающего': row['ФиоПодписывающего'],
                       'НазваниеОрганизации': row['НазваниеОрганизации'],
                       'Исполнитель': row['Исполнитель'], 'ДатаПодписанияДок': row['ДатаПодписанияДок'],
                       'Город': row['Город'],
                       'Год': row['Год']}
            doc.render(context)
            doc.save(f'{path_to_end_folder_scc}/{row["ФИОСлушателя"]}.docx')
        messagebox.showinfo('Dodger', 'Создание сертификатов успешно завершено!')

    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться сертификаты')


def generate_certificates():
    """
    Функция для создания удостоверений из шаблона и файла с данными
    :return:
    """
    try:
        # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
        reader = csv.DictReader(open(name_file_data_certificates), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)
        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_certificates)
            # Код для того чтобы операторы вводили номера без нулей при заполнении таблицы с данными
            number = ''
            if len(row['НомерДок']) == 2:
                number = '000' + row['НомерДок']
            elif len(row['НомерДок']) == 3:
                number = '00' + row['НомерДок']
            elif len(row['НомерДок']) == 4:
                number = '0' + row['НомерДок']
            else:
                number = row['НомерДок']
            context = {'ФИОСлушателя': row['ФИОСлушателя'], 'Фамилия': row['Фамилия'], 'НомерДок': number,
                       'ИмяОтчество': row['ИмяОтчество'], 'ФИОСлушателяРодПадеж': row['ФИОСлушателяРодПадеж'],
                       'region_genitive_case': row['region_genitive_case'],
                       'Оператор': row['Оператор'], 'РегионРодПадеж': row['РегионРодПадеж'],
                       'ВидПрограммы': row['ВидПрограммы'], 'ПодвидПрограммы': row['ПодвидПрограммы'],
                       'ФорматПрограммы': row['ФорматПрограммы'], 'НазваниеПрограммы': row['НазваниеПрограммы'],
                       'Профессия': row['Профессия'], 'Квалификация': row['Квалификация'],
                       'КатегорияРазряд': row['КатегорияРазряд'],
                       'Группа': row['Группа'], 'ФИОПреп': row['ФИОПреп'], 'КолЧасовОбуч': row['КолЧасовОбуч'],
                       'КолМесяцевОбуч': row['КолМесяцевОбуч'], 'БазаОбучения': row['БазаОбучения'],
                       'ТекстНачалоОбуч': row['ТекстНачалоОбуч'],
                       'ТекстОконОбуч': row['ТекстОконОбуч'], 'ЧислоНачалоОбуч': row['ЧислоНачалоОбуч'],
                       'ЧислоКонецОбуч': row['ЧислоКонецОбуч'],
                       'ДолжностьФиоРодПадеж': row['ДолжностьФиоРодПадеж'],
                       'ПолнаяСтоимость': row['ПолнаяСтоимость'], 'ПерваяЧастьОплаты': row['ПерваяЧастьОплаты'],
                       'ДатаПервойОплаты': row['ДатаПервойОплаты'],
                       'ВтораяЧастьОплаты': row['ВтораяЧастьОплаты'], 'ДатаВторойОплаты': row['ДатаВторойОплаты'],
                       'ТретьяЧастьОплаты': row['ТретьяЧастьОплаты'],
                       'ДатаТретьейОплаты': row['ДатаТретьейОплаты'],
                       'ДатаОкончанияДоговора': row['ДатаОкончанияДоговора'],
                       'ДолжностьПодписывающего': row['ДолжностьПодписывающего'],
                       'ФиоПодписывающего': row['ФиоПодписывающего'],
                       'НазваниеОрганизации': row['НазваниеОрганизации'],
                       'Исполнитель': row['Исполнитель'], 'ДатаПодписанияДок': row['ДатаПодписанияДок'],
                       'Город': row['Город'],
                       'Год': row['Год']}
            doc.render(context)
            doc.save(f'{path_to_end_folder_certificates}/{row["ФИОСлушателя"]}.docx')
        messagebox.showinfo('Dodger', 'Создание удостоверений  успешно завершено!')
    except NameError as e:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


# Создаем окно
window = Tk()
window.title('Dodger')
window.geometry('640x480')

# Создаем объект вкладок

tab_control = ttk.Notebook(window)

# Создаем вкладку свидетельства о повышении
tab_diplomas = ttk.Frame(tab_control)
tab_control.add(tab_diplomas, text='Создание свидетельств')
tab_control.pack(expand=1, fill='both')

# Добавляем виджеты на вкладку
# Создаем метку для описания назначения программы
lbl_hello = Label(tab_diplomas, text='Скрипт для создания свидетельств')
lbl_hello.grid(column=0, row=0, padx=10, pady=25)

# Создаем кнопку Выбрать шаблон

btn_template_certificate = Button(tab_diplomas, text='1) Выберите шаблон свидетельства', font=('Arial Bold', 20),
                                  command=select_file_template_diplomas, )
btn_template_certificate.grid(column=0, row=1, padx=10, pady=10)

# Создаем кнопку Выбрать файл с данными
btn_data_certificate = Button(tab_diplomas, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                              command=select_file_data_diplomas)
btn_data_certificate.grid(column=0, row=2, padx=10, pady=10)

# Создаем кнопку для выбора папки куда будут генерироваться файлы

btn_choose_end_folder_certificate = Button(tab_diplomas, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                           command=select_end_folder_diplomas)
btn_choose_end_folder_certificate.grid(column=0, row=3, padx=10, pady=10)

# Создаем кнопку для запуска функции генерации файлов

btn_create_files_certificate = Button(tab_diplomas, text='4) Создать свидетельства', font=('Arial Bold', 20),
                                      command=generate_diplomas)
btn_create_files_certificate.grid(column=0, row=4, padx=10, pady=10)

# Создаем вкладку для создания  сертификатов
tab_scc = ttk.Frame(tab_control)
tab_control.add(tab_scc, text='Создание сертификатов')

# Добавляем виджеты на вкладку
lbl_hello = Label(tab_scc, text='Скрипт для создания сертификатов')
lbl_hello.grid(column=0, row=0, padx=10, pady=25)

# Создаем кнопку Выбрать шаблон

btn_template_scc = Button(tab_scc, text='1) Выберите шаблон сертификата', font=('Arial Bold', 20),
                          command=select_file_template_scc, )
btn_template_scc.grid(column=0, row=1, padx=10, pady=10)

# Создаем кнопку Выбрать файл с данными
btn_data_scc = Button(tab_scc, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                      command=select_file_data_scc)
btn_data_scc.grid(column=0, row=2, padx=10, pady=10)

# Создаем кнопку для выбора папки куда будут генерироваться файлы

btn_choose_end_folder_scc = Button(tab_scc, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder_scc)
btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)

# Создаем кнопку для запуска функции генерации файлов

btn_create_files_scc = Button(tab_scc, text=' Создать сертификаты', font=('Arial Bold', 20),
                              command=generate_scc)
btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

# Создаем вкладку Создание удостоверений
tab_certificate = ttk.Frame(tab_control)
tab_control.add(tab_certificate, text='Создание удостоверений')

# Добавляем виджеты на вкладку
lbl_hello = Label(tab_certificate, text='Скрипт для создания удостоверений')
lbl_hello.grid(column=0, row=0, padx=10, pady=25)

# Создаем кнопку Выбрать шаблон

btn_template_scc = Button(tab_certificate, text='1) Выберите шаблон удостоверения', font=('Arial Bold', 20),
                          command=select_file_template_certificates)
btn_template_scc.grid(column=0, row=1, padx=10, pady=10)

# Создаем кнопку Выбрать файл с данными
btn_data_scc = Button(tab_certificate, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                      command=select_file_data_certificates)
btn_data_scc.grid(column=0, row=2, padx=10, pady=10)

# Создаем кнопку для выбора папки куда будут генерироваться файлы

btn_choose_end_folder_scc = Button(tab_certificate, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder_certificates)
btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)

# Создаем кнопку для запуска функции генерации файлов

btn_create_files_scc = Button(tab_certificate, text=' Создать удостоверения', font=('Arial Bold', 20),
                              command=generate_certificates)
btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

window.mainloop()
