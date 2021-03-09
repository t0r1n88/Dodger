from tkinter import *



# Создаем окно в котором будут отображаться элементы
window = Tk()

# Создаем виджет ярлыка
# background bg определяет цвет фона
greeting = Label(text='Во славу Омнисии!!!',
                 bg='#34A2FE',
                 width=50,
                 height=20,
                 font=40)
greeting.pack()
# Создаем поле для ввода текста
inp = Entry()
# Размещаем элемент
inp.pack()
# Создаем кнопку
button = Button(text='Создать',
                width=25,
                height=5,
                bg='blue',
                fg='yellow'
            )
# Размещаем кнопку
button.pack()

name = inp.get()
print(name)

window.mainloop()