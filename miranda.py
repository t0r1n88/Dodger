from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk
from pymorphy2 import MorphAnalyzer
import pandas as pd


def create_case_fio(dct_last, dct_first, dct_patr, changed, lst_cases):
    """
    Функция для объединения склоняемых слов в ФИО. Сейчас необходимы слова в родительном падеже
    :param dct_last: Словарь с фамилиями
    :param dct_first: Словарь с именами
    :param dct_patr: Словарь с отчествами
    :param changed: Булев .True если слово можно просклонять, False если нет
    :param lst_cases: список падежей
    :return: словарь с просклонянеными ФИО
    """
    # Создаем словарь где ключ это падеж а значение ФИО в соответствующем падеже
    dct_fio = {case: '' for case in lst_cases}
    dct_fio['change'] = changed
    # Добавляем данные в словарь
    for case in lst_cases:
        dct_fio[case] = f'{dct_last[case]} {dct_first[case]} {dct_patr[case]}'.title()

    return dct_fio


def parse_case(word, gender, case, tag_fio, morph):
    """
    Функция для проверки возможности склонения слова по роду и падежу
    :param word: проверяемое слово
    :param gender: проверяемый пол
    :param case: проверяемый падеж
    :param morph: анализатор
    :return: Возвращает слово в неизменненом виде если нельзя просклонять по падежу, если можно то измененное
    """
    # Парсим слово получаем все возможные лексемы
    word_parsed = morph.parse(word)
    # Перебираем полученные разборы на предмет совпадений
    for par in word_parsed:
        if (tag_fio in par.tag) and (gender in par.tag):
            return par.inflect({gender, case}).word, 1

    return word, 0


def create_cases(lastname, firstname, patronymic, gender, lst_cases, morph):
    """
    Функция для склонения слов по падежам.Вероятно можно было бы сделать поизящней
    :param lastname: Фамилия
    :param firstname: Имя
    :param patronymic: Отчество
    :param gender: Пол
    :param lst_cases: Список падежей
    :return: 3 словаря в каждом из которых слово просклонено по 6 падежам и признак того что слово удалось просклонять
    """
    # Создаем теги наличие которых будет говорить о возможности просклонять слово
    tag_lastname = 'Surn'
    tag_firstname = 'Name'
    tag_patr = 'Patr'
    # Создаем словари где ключом будет падеж а значение слово в соответсвтующем падеже , хотя правильнее было бы использовать словари, где ключом был бы падеж
    # Создаем с помощью генераторов словарей
    dct_lastname = {case: '' for case in lst_cases}
    dct_firstname = {case: '' for case in lst_cases}
    dct_patronymic = {case: '' for case in lst_cases}
    lst_changed = []
    # Перебираем список падежей и при каждой итерации добавляем в словарь по соответствующему ключу слово в текущем падеже
    for case in lst_cases:
        dct_lastname[case], changed_lastname = parse_case(lastname, gender, case, tag_lastname, morph)
        dct_firstname[case], changed_firstname = parse_case(firstname, gender, case, tag_firstname, morph)
        dct_patronymic[case], changed_patronymic = parse_case(patronymic, gender, case, tag_patr, morph)
        lst_changed.extend([changed_lastname, changed_firstname, changed_patronymic])

    return dct_lastname, dct_firstname, dct_patronymic, all(lst_changed)


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
    Функция для выбора файла с данными на основе которых будет генерироваться договор  и генерация падежей фио
    :return: Путь к файлу с данными и словарь с просклоняемыми ФИО
    """
    global name_file_data_contracts
    # Получаем путь к файлу
    name_file_data_contracts = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))
    # Загружаем полученный датафрейм беря из него только колонку ФиоСлушателя
    base_df = pd.read_csv(name_file_data_contracts, delimiter=';', encoding='cp1251', usecols=['ФиоСлушателя', 'Пол'])
    # Начинаем обработку
    # Создаем объект анализатор
    morph = MorphAnalyzer()
    # Список падежей в которые нужно просклонять
    lst_cases = ['nomn', 'gent', 'datv', 'accs', 'ablt', 'loct']
    # Счетчик строк для того чтобы можно было узнать на какой строке возникла проблема
    counter_rows = 2
    # Создаем словарь в котором будем хранить просклоненные фио. Придется сделать его глобальным
    global case_fio_dct
    case_fio_dct = {}

    for row in base_df.itertuples():
        try:
            # Создаем список из строки,делим по пробелам
            lastname, firstname, patronymic = row[1].split()
            gender = 'masc' if row[2] == 1 else 'femn'
            # Склоняем слова
            dct_lastname, dct_firstname, dct_patronymic, changed = create_cases(lastname, firstname, patronymic, gender,
                                                                                lst_cases,
                                                                                morph)
            value_to_table = create_case_fio(dct_lastname, dct_firstname, dct_patronymic, changed, lst_cases)
            # Добавляем обработанное ФИО в словарь.
            value_to_table['changed'] = changed

            temp_dct = {}
            temp_dct[f'{value_to_table["nomn"]}'] = value_to_table

            case_fio_dct.update(temp_dct)
            counter_rows += 1

        except ValueError:
            messagebox.showerror(
                message=f'Произошла ошибка.Проверьте строку {counter_rows} в файле. Отсутствует часть ФИО')


def select_end_folder_contracts():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_contracts
    path_to_end_folder_contracts = filedialog.askdirectory()


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

            context = {'ФиоСлушателя': row['ФиоСлушателя'],
                       'ФиоСлушателяРодПадеж': case_fio_dct[row['ФиоСлушателя']]['gent'],
                       'НомерДоговора': row['НомерДоговора'],
                       'ДатаПодписанияДоговора': row['ДатаПодписанияДоговора'],
                       'ДолжностьФиоРодительныйПадеж': row['ДолжностьФиоРодительныйПадеж'],
                       'Программа': row['Программа'], 'СрокВМесяцах': row['СрокВМесяцах'],
                       'Профессия': row['Профессия'], 'СрокВЧасах': row['СрокВЧасах'],
                       'ДатаНачалаЗанятий': row['ДатаНачалаЗанятий'],
                       'НачалоОбучения': row['НачалоОбучения'],
                       'КонецОбучения': row['КонецОбучения'], 'ПолнаяСтоимость': row['ПолнаяСтоимость'],
                       'ПерваяЧастьОплаты': row['ПерваяЧастьОплаты'],
                       'ДатаПервойОплаты': row['ДатаПервойОплаты'],
                       'ВтораяЧастьОплаты': row['ВтораяЧастьОплаты'],
                       'ДатаВторойОплаты': row['ДатаВторойОплаты'],
                       'ТретьяЧастьОплаты': row['ТретьяЧастьОплаты'],
                       'ДатаТретьейОплаты': row['ДатаТретьейОплаты'],
                       'ДатаОкончанияДоговора': row['ДатаОкончанияДоговора'],
                       'ДолжностьПодписывающего': row['ДолжностьПодписывающего'],
                       'ФиоПодписывающего': row['ФиоПодписывающего'],
                       'ДатаПодписиДоговора': row['ДатаПодписиДоговора'], 'ДатаРождения': row['ДатаРождения'],
                       'СерияПаспорта': row['СерияПаспорта'], 'НомерПаспорта': row['НомерПаспорта'],
                       'ДатаВыдачиПаспорта': row['ДатаВыдачиПаспорта'], 'Выдан': row['Выдан'],
                       'АдресРегистрации': row['АдресРегистрации'], 'Снилс': row['Снилс'],
                       'КонтактныйТелефон': row['КонтактныйТелефон']}
            doc.render(context)
            if case_fio_dct[row['ФиоСлушателя']]['changed']:
                doc.save(f'{path_to_end_folder_contracts}/{row["ФиоСлушателя"]}.docx')
            else:
                doc.save(f'{path_to_end_folder_contracts}/Проверьте склонение ФИО {row["ФиоСлушателя"]}.docx')

        messagebox.showinfo('Miranda', 'Создание договоров успешно завершено!')
    except NameError as e:
        messagebox.showinfo('Miranda', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')

def generate_order_enroll():
    """
    Функция для создания сертификатов

    """
    try:
        reader = csv.DictReader(open(name_file_data_order_enroll), delimiter=';')
        # Конвертируем объект reader в список словарей
        data = list(reader)

        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_order_enroll)

            context = {'dative_case_lastname': row['dative_case_lastname'],
                       'dative_case_firstname': row['dative_case_firstname'],
                       'time': row['time'],
                       'category_program': row['category_program'], 'format_program': row['format_program'],
                       'name_program': row['name_program'],
                       'hour': row['hour'],
                       'chief_copp': row['chief_copp'], 'city': row['city'], 'year': row['year']
                       }
            doc.render(context)
            doc.save(f'{path_to_end_folder_order_enroll}/{row["dative_case_lastname"]} {row["dative_case_firstname"]}.docx')
        messagebox.showinfo('Dodger', 'Создание сертификатов успешно завершено!')

    except NameError:
        messagebox.showinfo('Dodger', 'Выберите шаблон,файл с данными и папку куда будут генерироваться сертификаты')

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



def select_file_template_order_enroll():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_order_enroll
    name_file_template_order_enroll = filedialog.askopenfilename(filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_order_enroll():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться
    :return: Путь к файлу с данными
    """
    global name_file_data_order_enroll
    name_file_data_order_enroll = filedialog.askopenfilename(filetypes=(('Csv files', '*.csv'), ('all files', '*.*')))


def select_end_folder_order_enroll():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_order_enroll
    path_to_end_folder_order_enroll = filedialog.askdirectory()


# Создаем окно
if __name__ == '__main__':
    window = Tk()
    window.title('Miranda')
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




    # Создаем вкладку для создания приказов о зачислении
    tab_order_enroll = ttk.Frame(tab_control)
    tab_control.add(tab_order_enroll, text='Создание приказов о зачислении')

    # Добавляем виджеты на вкладку
    lbl_hello = Label(tab_order_enroll, text='Скрипт для создания приказов о зачислении')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон

    btn_template_scc = Button(tab_order_enroll, text='1) Выберите шаблон приказа', font=('Arial Bold', 20),
                              command=select_file_template_order_enroll, )
    btn_template_scc.grid(column=0, row=1, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_scc = Button(tab_order_enroll, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                          command=select_file_data_order_enroll)
    btn_data_scc.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_scc = Button(tab_order_enroll, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_order_enroll)
    btn_choose_end_folder_scc.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для запуска функции генерации файлов

    btn_create_files_scc = Button(tab_order_enroll, text=' Создать приказы', font=('Arial Bold', 20),
                                  command=generate_order_enroll)
    btn_create_files_scc.grid(column=0, row=4, padx=10, pady=10)

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
