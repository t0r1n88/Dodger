import pandas as pd

df_big = pd.read_csv('resources/1949.csv')
df_smaal = pd.read_csv('resources/961.csv')
big_set = set(df_big['email'])
small_set = set(df_smaal['email'])
print(len(big_set))
print(len(small_set))
diff_set = big_set-small_set
print(diff_set)
