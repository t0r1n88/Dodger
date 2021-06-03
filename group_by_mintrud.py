import pandas as pd

df =pd.read_excel('resources/mintrud_df.xlsx')


df_okvad = df.groupby('Основной вид экономической деятельности (по ОКВЭД 2)')['Основной вид экономической деятельности (по ОКВЭД 2)'].count()
# df_okvad.columns = ['Основной вид экономической деятельности (по ОКВЭД 2)','Количество предприятий']
# print(df_okvad.columns)
# print(df_okvad.head())
print(df_okvad.columns)

df_okvad.sort_values(ascending=False)
df_okvad.to_excel('group_okvad_simple.xlsx')