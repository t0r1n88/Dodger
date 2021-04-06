from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pytest
from dodger import *


def test_example():
    assert 'Lindy' in example()

def test_create_window():
    """
    Дано: функция создающая окно ткинтер
    Когда : появляется окно программы
    Тогда : заголовок программы должен содеражать название. Пока это самая простая проверка
    """
    window = create_window()
    title = window.title
    assert 'Dodger' in title()
