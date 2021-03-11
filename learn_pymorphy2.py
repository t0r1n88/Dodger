from pymorphy2 import MorphAnalyzer

# Создаем объект анализатор
morph = MorphAnalyzer()

# Слово для примера
word = 'Шиша'
lst_case = ['nomn', 'gent', 'datv', 'accs', 'ablt', 'loct']

# Анализируем слово
word_parsed =morph.parse(word)[0]
# print(word_parsed.tag.gender)
# print(word_parsed.lexeme)
# for case in lst_case:
#     print(word_parsed)
#     print(word_parsed.inflect({case}).word)
print(len(word_parsed.lexeme))
for par in word_parsed.lexeme:

    if par.tag.gender == 'femn':
        print(par)