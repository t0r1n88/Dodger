from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk


class ProgramWindow:
    """
    Класс для создания окна программы и элементов
    """

    def __init__(self, title='Dodger',geometry='640x480+700+300'):
        # Создаем объект окна
        self.window = Tk()
        # Присваиваем имя окну
        self.window.title = title
        # Указываем размеры окна
        self.window.geometry(geometry)
        # Создаем объект для вкладок
        self.tab_control = ttk.Notebook(self.window)

    def add_tab(self, title_tab, name_button):
        """
        Функция для добавления вкладок
        :param title_tab: Название вкладки
        """
        # Создаем фрейм для определенной вкладки
        frame = ttk.Frame(self.tab_control)
        # Привязываем фрейм к вкладке
        self.tab_control.add(frame, text=title_tab)
        # Размещаем кнопку
        button = ttk.Button(frame, text=name_button)
        button.grid(column=1, row=1,padx=10, pady=25)
        self.tab_control.pack(expand=1, fill='both')

    def run(self):
        self.window.mainloop()


def create_window():
    window = Tk()
    window.title('Dodger')
    window.geometry('640x480+700+300')

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку свидетельства о повышении
    tab_diplomas = ttk.Frame(tab_control)
    tab_control.add(tab_diplomas, text='Создание свидетельств')
    tab_control.pack(expand=1, fill='both')

    return window


if __name__ == '__main__':
    window = ProgramWindow()
    window.add_tab('Создание свидетельств', 'Создать свидетельства')
    window.add_tab('Создание сертификатов', 'Создать сертификаты')
    print(dir(window.tab_control))


    window.run()
