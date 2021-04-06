from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import csv
from docxtpl import DocxTemplate
from tkinter import ttk

def example():
    return 'Lindy Booth'


def create_window():
    window = Tk()
    window.title('Dodger')
    window.geometry('640x480')

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку свидетельства о повышении
    tab_diplomas = ttk.Frame(tab_control)
    tab_control.add(tab_diplomas, text='Создание свидетельств')
    tab_control.pack(expand=1, fill='both')
    return window
if __name__ == '__main__':
    window = create_window()
    window.mainloop()