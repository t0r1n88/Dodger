from twocaptcha import TwoCaptcha
import time
import pyperclip

"""
Скрипт для работы с сервисом по расшифровке капч
"""
API_KEY = open('API_captcha.txt').read()

# Поменять параметры
# Создаем объект
solver = TwoCaptcha(API_KEY,server='rucaptcha.com')
result = None
timeout = 40
# получаем id капчу отправленной на сервер
# Пока не получим результат от сервера будем постепенно увеличивать время ожидания обработки
while not result:
    try:
        id = solver.send(file='resources/captcha.jpg')
        time.sleep(timeout)
        result = solver.get_result(id)
        pyperclip.copy(result)
    except Exception as e:
        timeout += 5
        continue

exit()



