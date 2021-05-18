# # Создаем класс
# class Car:
#     # Создаем атрибуты класса
#
#     car_count = 0
#     def __init__(self):
#         Car.car_count += 1
#         print(Car.)
#     # Создаем методы класса
#     def start(self):
#         print('Заводим двигатель')
#
#     def stop(self):
#         print('Остановить двигатель')
#
# # Создаем 2 объекта класса
# car_a = Car()
# car_b = Car()
#
# print(dir(car_b))
import tkinter as tk

root = tk.Tk()
root.title('Главная программа')

frame_start = tk.Frame(root)

frame_start.pack()

button1 = tk.Button(master=frame_start, text='Интегрирование', bg='green', fg='red', width=39,
                    command=lambda: change_frame(1))  # я знаю что здесь будет , command = Название функции
button2 = tk.Button(master=frame_start, text='Решение нелинейного уравнения /\nдифференциального уравнения', bg='pink',
                    fg='cyan', width=39, command=lambda: change_frame(2))
button3 = tk.Button(master=frame_start, text='Построение графика \nпараметрически заданной функции', bg='purple',
                    fg='brown', width=39, command=lambda: change_frame(3))
button4 = tk.Button(master=frame_start, text='Конструктор блок-схем', bg='cyan', fg='orange', width=39,
                    command=lambda: change_frame(4))

button1.pack()
button2.pack()
button3.pack()
button4.pack()

# Интегрирование
frame_1 = tk.Frame(root)
label_1 = tk.Label(master=frame_1, text='Интегрирование')
button_1 = tk.Button(master=frame_1, text='Назад', command=lambda: back(1))
label_1.pack()
button_1.pack()

# Решение нелинейного уравнения /\nдифференциального уравнения
frame_2 = tk.Frame(root)
label_2 = tk.Label(master=frame_2, text='Решение нелинейного уравнения /\nдифференциального уравнения')
button_2 = tk.Button(master=frame_2, text='Назад', command=lambda: back(2))
label_2.pack()
button_2.pack()

# Построение графика \nпараметрически заданной функции
frame_3 = tk.Frame(root)
label_3 = tk.Label(master=frame_3, text='Построение графика \nпараметрически заданной функции')
button_3 = tk.Button(master=frame_3, text='Назад', command=lambda: back(3))
label_3.pack()
button_3.pack()

# Конструктор блок-схем
frame_4 = tk.Frame(root)
label_4 = tk.Label(master=frame_4, text='Конструктор блок-схем')
button_4 = tk.Button(master=frame_4, text='Назад', command=lambda: back(4))
label_4.pack()
button_4.pack()


# поменять фрейм
def change_frame(num):
    frame_start.forget()
    eval('frame_' + str(num)).pack()


# вернутся к стартовому фрейму
def back(num):
    eval('frame_' + str(num)).forget()
    frame_start.pack()


root.mainloop()