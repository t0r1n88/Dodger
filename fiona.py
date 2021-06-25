from twocaptcha import TwoCaptcha
import time
import pyperclip

"""
Скрипт для работы с сервисом по расшифровке капч
"""
API_KEY = open('API_captcha.txt').read()

# Поменять параметры
# Создаем объект
solver = TwoCaptcha(API_KEY, server='rucaptcha.com')
result = None
timeout = 40
count = 0
# получаем id капчу отправленной на сервер
# Пока не получим результат от сервера будем постепенно увеличивать время ожидания обработки
while not result:
    try:
        id = solver.send(file='resources/captcha.jpg')
        print('Капча отправлена')
        count += 1
        if count == 5:
            exit()
        time.sleep(timeout)
        result = solver.get_result(id)
        if not result:
            continue
        print(result)
        pyperclip.copy(result)
    except Exception as e:
        timeout += 5
        continue

exit()
