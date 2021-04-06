from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pytest

from dodger import *

def test_create_window():
    """
    Дано: функция создающая окно ткинтер
    Когда : появляется окно программы
    Тогда : заголовок программы должен содеражать название. Пока это самая простая проверка
    """
    window = create_window()
    assert 'Dodger' in window.title()

def test_tabs_exists():
    """
    Дано: окно ткинтер со вкладками
    Когда: программа запущена
    Тогда: Все вкладки должны быть рабочими
    """

