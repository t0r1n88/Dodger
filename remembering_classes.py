# Создаем класс
class Car:
    # Создаем атрибуты класса

    car_count = 0
    def __init__(self):
        Car.car_count += 1
        print(Car.)
    # Создаем методы класса
    def start(self):
        print('Заводим двигатель')

    def stop(self):
        print('Остановить двигатель')

# Создаем 2 объекта класса
car_a = Car()
car_b = Car()

print(dir(car_b))