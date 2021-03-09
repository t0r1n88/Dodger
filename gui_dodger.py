from tkinter import *
from tkinter import filedialog
from tkinter import messagebox


# Функция выбора шаблона

def select_file_template():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    name_file_template = filedialog.askopenfilename(filetypes=(('Word files', '*.docx'), ('all files', '*.*')))
    print(name_file_template)
    return name_file_template


def select_file_data():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    name_file_data = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))
    print(name_file_data)

def select_end_folder():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    path_to_end_folder = filedialog.askdirectory()
    print(path_to_end_folder)

def generate_files():
    """
    Функция для создания файлов из шаблона и файла с данными
    :return:
    """
    messagebox.showinfo('Создано','Во славу Омниссии!!!')


# Создаем окно
window = Tk()
window.title('Dodger')
window.geometry('640x480')

# Создаем метку для описания назначения программы
lbl_hello = Label(window, text='Программа для генерации документов из шаблонов')
lbl_hello.grid(column=0, row=0)

# Создаем кнопку Выбрать шаблон

btn_template = Button(window, text='Выберите шаблон документа', font=('Arial Bold', 20),
                      command=select_file_template, )
btn_template.grid(column=0, row=1)

# Создаем кнопку Выбрать файл с данными
btn_data = Button(window, text='Выберите файл с данными', font=('Arial Bold', 20),
                  command=select_file_data)
btn_data.grid(column=0, row=2)

# Создаем кнопку для выбора папки куда будут генерироваться файлы

btn_choose_end_folder = Button(window, text='Выберите конечную папку', font=('Arial Bold', 20),
                               command=select_end_folder)
btn_choose_end_folder.grid(column=0,row=3)

# Создаем кнопку для запуска функции генерации файлов

btn_create_files = Button(window,text=' Создать документы', font=('Arial Bold',20),
                          command=generate_files)
btn_create_files.grid(column=0, row=4)

window.mainloop()
