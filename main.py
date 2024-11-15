import pandas as pd

df = pd.read_excel("Перечень сформированных угроз мой.xlsx", sheet_name='Перечень угроз')
protective_measures = df['Меры защиты']

df_result = pd.DataFrame(columns=['Код', 'Меры защиты', 'Баллы'])
dict_measures = {}

for str_item in protective_measures:
    str_arr = str_item.replace("\n", '').split("_x000d_")
    for i in str_arr:
        if dict_measures.get(i):
            dict_measures[i] = dict_measures.get(i) + 1
        else:
            dict_measures[i] = 1

for i in dict_measures:
    cod = i[0: i.find(' ')]
    txt = i[i.find(' '): len(i)]
    count = dict_measures[i]
    df_result.loc[len(df_result.index)] = [cod, txt, count]

df_result = df_result.sort_values(ascending=False, by= 'Баллы')
df_result.to_excel('result.xlsx')

print(df_result)





