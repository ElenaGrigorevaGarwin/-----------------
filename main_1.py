import os
import pandas as pd

df = pd.read_excel('Факт отгрузок.xlsx')

df.columns = df.columns[:1].tolist() + df.columns[1:].map(lambda x: x.strftime('%d.%m.%Y')).tolist()

df_transposed = df.T  # Транспонирование таблицы
df_transposed.columns = df_transposed.iloc[0]  # Сделать первую строку заголовком
df_transposed = df_transposed.iloc[1:]  # Удалить первую строку после использования как заголовка

# Сброс индекса и перевод его в обычный столбец, а также присвоение имени столбцу "Дата"
df_transposed.reset_index(inplace=True)
df_transposed.rename(columns={'index': 'Дата'}, inplace=True)

# Заполнение отсутствующих данных нулями
df_transposed.fillna(0, inplace=True)

df_transposed.to_excel('Факт отгрузок.xlsx', index=False)  # запись итогового файла без индекса