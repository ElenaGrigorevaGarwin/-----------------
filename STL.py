
import pandas as pd
import statsmodels.api as sm

df = pd.read_excel('Факт отгрузок.xlsx')
df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')
df.set_index('Дата', inplace=True)
predictions = pd.DataFrame()

for department in df.columns:
  sales_data = df[department]
  result = sm.tsa.STL(sales_data, seasonal=13).fit()
  forecast_steps = 12
  forecast = result.trend[-1] + result.seasonal[-12:]
  predictions[department] = forecast.values

start_date = df.index[-1] + pd.DateOffset(months=1)
predictions.index = pd.date_range(start=start_date, periods=12, freq='M')
# Сохраняем в таблицу
predictions.to_excel("result/прогнозы_STL.xlsx")