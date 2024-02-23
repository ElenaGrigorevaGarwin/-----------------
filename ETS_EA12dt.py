import pandas as pd
import numpy as np
import statsmodels.api as sm

df = pd.read_excel('Факт отгрузок.xlsx')

# Преобразуйте столбец 'Дата' в datetime и установите его как индекс
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')
df.set_index('Дата', inplace=True)

# Установите частоту 'M' (ежемесячно) явно
df.index = pd.date_range(df.index[0], periods=len(df), freq='MS')
# Создайте пустой DataFrame для прогнозов
predictions = pd.DataFrame()
for department in df.columns:
    sales_data = df[department]
    ets_model = sm.tsa.statespace.ExponentialSmoothing(sales_data,
                                                       trend='add',  # 'add'  и 'multiplicative'
                                                       seasonal=12, # 4 и 12   
                                                       initialization_method= 'estimated',   # метод инициализации - 'estimated', 'heuristic', 'concentrated'
                                                       damped_trend=True).fit()     # затухающий тренд - сделать прогноз и без него
    forecast_values = ets_model.forecast(12)
    predictions[department] = forecast_values

# Добавление индекса (месяц и год) в DataFrame
start_date = df.index[-1] + pd.DateOffset(months=1)
predictions.index = pd.date_range(start=start_date, periods=12, freq='MS')
# Сохранение прогнозов в Excel
predictions.to_excel("result/Прогнозы_ETS_EA12dt.xlsx")