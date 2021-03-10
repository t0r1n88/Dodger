from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk


# Функции выбора шаблона,файла с данными и конеченой папки для генерации удостоверений

def select_file_template_certificates():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_certificates
    name_file_template_certificates = filedialog.askopenfilename(filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


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
            if len(row['number']) == 2:
                number = '000' + row['number']
            elif len(row['number']) == 3:
                number = '00' + row['number']
            elif len(row['number']) == 4:
                number = '0' + row['number']
            else:
                number = row['number']
            context = {'lastname': row['lastname'], 'firstname': row['firstname'], 'number': number,
                       'region_genitive_case': row['region_genitive_case'],
                       'educator': row['educator'], 'type_program': row['type_program'],
                       'profession': row['profession'], 'date_expiry': row['date_expiry'],
                       'date_issue': row['date_issue'],
                       'qualification': row['qualification'],
                       'category': row['category'], 'name_prep': row['name_prep'], 'name_dir': row['name_dir'],
                       'hour': row['hour'], 'base': row['base'], 'begin': row['begin'], 'end': row['end']}
            doc.render(context)
            doc.save(f'{path_to_end_folder_certificates}/{row["lastname"]} {row["firstname"]}.docx')
        messagebox.showinfo('Dodger', 'Создание удостоверений успешно завершено!')
    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')

def generate_scc():
    """
    Функция для создания сертификатов
    :return:
    """
    try:
        reader = csv.DictReader(open(name_file_data_scc), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)

        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_scc)

            context = {'dative_case_lastname': row['dative_case_lastname'],
                       'dative_case_firstname': row['dative_case_firstname'],
                       'time': row['time'],
                       'category_program': row['category_program'], 'format_program': row['format_program'],
                       'name_program': row['name_program'],
                       'hour': row['hour'],
                       'chief_copp': row['chief_copp'], 'city': row['city'], 'year': row['year']
                       }
            doc.render(context)
            doc.save(f'{path_to_end_folder_scc}/{row["dative_case_lastname"]} {row["dative_case_firstname"]}.docx')
        messagebox.showinfo('Dodger', 'Создание сертификатов успешно завершено!')

    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться сертификаты')





# Создаем окно
window = Tk()
window.title('Dodger')
window.geometry('640x480')

# Создаем объект вкладок

tab_control = ttk.Notebook(window)


# Создаем вкладку свидетельства о повышении
tab_certificate = ttk.Frame(tab_control)
tab_control.add(tab_certificate, text='Создание свидетельств')
tab_control.pack(expand=1, fill='both')


# Добавляем виджеты на вкладку
# Создаем метку для описания назначения программы
lbl_hello = Label(tab_certificate, text='Скрипт для создания свидетельств')
lbl_hello.grid(column=0, row=0, padx=10, pady=25)

# Создаем кнопку Выбрать шаблон

btn_template_certificate = Button(tab_certificate, text='1) Выберите шаблон свидетельства', font=('Arial Bold', 20),
                                  command=select_file_template_certificates, )
btn_template_certificate.grid(column=0, row=1, padx=10, pady=10)

# Создаем кнопку Выбрать файл с данными
btn_data_certificate = Button(tab_certificate, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                              command=select_file_data_certificates)
btn_data_certificate.grid(column=0, row=2, padx=10, pady=10)

# Создаем кнопку для выбора папки куда будут генерироваться файлы

btn_choose_end_folder_certificate = Button(tab_certificate, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                           command=select_end_folder_certificates)
btn_choose_end_folder_certificate.grid(column=0, row=3, padx=10, pady=10)

# Создаем кнопку для запуска функции генерации файлов

btn_create_files_certificate = Button(tab_certificate, text='4) Создать свидетельства', font=('Arial Bold', 20),
                                      command=generate_certificates)
btn_create_files_certificate.grid(column=0, row=4, padx=10, pady=10)





# Создаем вкладку сертификатов
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


# Создаем диплом о профессиональной переподготовке
tab_diploma = ttk.Frame(tab_control)
tab_control.add(tab_diploma, text='Создание удостоверений')

window.mainloop()
