import pandas as pd

df = pd.read_excel('data/Список организаций (Название,ИНН).xlsx',dtype={'Наименование':str,'ИНН':str})
# Удаляем организации для которых отсутствует ИНН
df.dropna(inplace=True)

# Добавляем ноль
df['ИНН'] = df['ИНН'].apply(lambda x:'0'+x if len(x) == 9 else x)
# Созраняем в csv
df.to_csv('inn.csv',encoding='cp1251')