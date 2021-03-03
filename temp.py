from tkinter import *
from tkinter import filedialog
# Создаем глобальный объект окна
root = Tk()

# Настройки окна

def foo():
    global myfile
    myfile = filedialog.askopenfilename()
    print((myfile))

root['bg'] = '#fafafa'
# Заголовок окна
root.title('Для создания заявлений')
root.wm_attributes('-alpha',1)
root.geometry('640x480')

root.resizable(width=False, height=False)


# Создадим холст на котором будем отрисовывать элементы
canvas = Canvas(root, height=300, width=250)
canvas.pack()

# Создаем фрейм
frame = Frame(root,bg='green')
# frame.place(relx=0.15, rely=0.15, relwidth=0.7, relheight=0.7)
frame.place(relx=0.15, rely=0.15, relwidth=0.7, relheight=0.7)

# Создаем элементы

title = Label(frame,text='Во славу Omnissiah!!!', bg='cyan', font =40)
# Располагаем элемент на фрейме
title.pack()
btn = Button(frame, text='Кнопка для дела',command=foo, bg='yellow')
btn.pack()
root.mainloop()
