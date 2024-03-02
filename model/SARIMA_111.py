# SARIMA 111
import pandas as pd
import numpy as np
import statsmodels.api as sm

# Чтение данных
df = pd.read_excel('Факт отгрузок.xlsx')

# Преобразование столбца 'Дата' в datetime
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

# Установка 'Дата' в качестве индекса
df.set_index('Дата', inplace=True)

# Построение модели SARIMA и создание прогнозов
predictions = pd.DataFrame()

for department in df.columns:
    sales_data = df[department]

    # Построение SARIMA модели
    model = sm.tsa.SARIMAX(sales_data, order=(1, 1, 1), seasonal_order=(1, 1, 1, 12))
    results = model.fit()

    # Прогноз на следующий год (12 месяцев)
    forecast = results.get_forecast(steps=12)
    forecast_values = forecast.predicted_mean

    # Добавление прогнозов в DataFrame
    predictions[department] = forecast_values

# Добавление индекса (месяц и год) в DataFrame
start_date = df.index[-1] + pd.DateOffset(months=1)
predictions.index = pd.date_range(start=start_date, periods=12, freq='MS')

# Сохранение прогнозов в Excel
predictions.to_excel("result/прогнозы_SARIMA_111.xlsx")