from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk


# Функция выбора шаблона

def select_file_template():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template
    name_file_template = filedialog.askopenfilename(filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data
    name_file_data = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()


def generate_files():
    """
    Функция для создания файлов из шаблона и файла с данными
    :return:
    """
    try:
        # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
        reader = csv.DictReader(open(name_file_data), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)
        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template)
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
                       'profession': row['profession'], 'date_expiry': row['date_expiry'],
                       'date_issue': row['date_issue'],
                       'qualification': row['qualification'],
                       'category': row['category'], 'name_prep': row['name_prep'], 'name_dir': row['name_dir'],
                       'hour': row['hour'], 'base': row['base'], 'begin': row['begin'], 'end': row['end']}
            doc.render(context)
            doc.save(f'{path_to_end_folder}/{row["lastname"]} {row["firstname"]}.docx')
        messagebox.showinfo('Dodger', 'Создание файлов успешно завершено!')
    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')





# Создаем окно
window = Tk()
window.title('Dodger')
window.geometry('640x480')

# Создаем объект вкладок

tab_control = ttk.Notebook(window)
# Создаем вкладку свидетельства о повышении
tab_scc = ttk.Frame(tab_control)
tab_control.add(tab_scc, text='Создание свидетельств')
tab_control.pack(expand=1, fill='both')

# Создаем вкладку удостоверний о повышени  квалификации
tab_certificate = ttk.Frame(tab_control)
tab_control.add(tab_certificate, text='Создание удостоверений')

# Создаем диплом о профессиональной переподготовке
tab_diploma = ttk.Frame(tab_control)
tab_control.add(tab_diploma, text='Создание дипломов проф. переподготовки')
# Создаем метку для описания назначения программы
lbl_hello = Label(tab_scc, text='Создание свидетельств')
lbl_hello.grid(column=0, row=0, padx=10, pady=25)

# Создаем кнопку Выбрать шаблон

btn_template = Button(tab_scc, text='Выберите шаблон документа', font=('Arial Bold', 20),
                      command=select_file_template, )
btn_template.grid(column=0, row=1, padx=10, pady=10)

# Создаем кнопку Выбрать файл с данными
btn_data = Button(tab_scc, text='Выберите файл с данными', font=('Arial Bold', 20),
                  command=select_file_data)
btn_data.grid(column=0, row=2, padx=10, pady=10)

# Создаем кнопку для выбора папки куда будут генерироваться файлы

btn_choose_end_folder = Button(tab_scc, text='Выберите конечную папку', font=('Arial Bold', 20),
                               command=select_end_folder)
btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

# Создаем кнопку для запуска функции генерации файлов

btn_create_files = Button(tab_scc, text=' Создать документы', font=('Arial Bold', 20),
                          command=generate_files)
btn_create_files.grid(column=0, row=4, padx=10, pady=10)

window.mainloop()
