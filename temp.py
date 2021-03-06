c = """Основной (по коду ОКВЭД ред.2): <a title="Эта группировка включает:
- установку электротехнических систем во всех видах зданий и сооружений гражданского строительства;
- монтаж электропроводки и электроарматуры, телекоммуникаций, компьютерной сети и проводки кабельного телевидения, включая оптоволоконные линии связи, антенн всех типов, включая спутниковые антенны, осветительных систем, пожарной сигнализации, систем охранной сигнализации, уличного освещения и иного электрооборудования на автомобильных дорогах, энергообеспечения наземного электротранспорта и электротехнического сигнального оборудования, освещения взлетно-посадочных полос аэропортов и космодромов, электрических коллекторов солнечной энергии;
- выполнение работ по подводке электросетей для подключения электроприборов, кодовых замков, домофонов и прочего оборудования, включая плинтусное отопление
Эта группировка также включает:
- приспособление систем электрообеспечения на объектах культурного наследия
Эта группировка не включает:
- приспособление инженерных систем и оборудования на объектах культурного наследия, см. 43.29" href="/list?okved2=43.21" winautomationvisibilitylandmark="true">43.21 - Производство электромонтажных работ
"""
lst = (c.split('\n'))

# Пример генератора словарей. И уже потом с помощью get забирать данные оттуда, и не надо будет регулярки писать для каждой колонки
dict_data = {row.split(':')[0]:row.split(':')[1].strip() for row in lst }

print(dict_data)