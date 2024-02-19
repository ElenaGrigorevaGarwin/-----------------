#Holt-Winters выбор параметров ('mul', 'mul')
import pandas as pd
from statsmodels.tsa.holtwinters import ExponentialSmoothing as HWES

df = pd.read_excel('Факт отгрузок.xlsx')
df['Дата'] = pd.to_datetime(df['Дата'])
sales_forecast = pd.DataFrame()

for column in df.columns[1:]:
    series = pd.to_numeric(df[column], errors='coerce')

    # При необходимости исправьте нулевые значения
    series = series.replace(0, 1)

    model = HWES(series, seasonal_periods=12, trend='mul', seasonal='mul')
    fitted = model.fit()

    # Создание временного ряда прогнозов
    forecast_values = fitted.forecast(steps=12)

    # Создание DataFrame с прогнозами для текущего отдела
    forecast_df = pd.DataFrame({
        'Прогноз': forecast_values.values,
    }, index=pd.date_range(start=df['Дата'].max() + pd.DateOffset(months=1), periods=12, freq='M'))

    # Добавление столбца с прогнозами в общий DataFrame
    sales_forecast[column] = forecast_df['Прогноз']

# Сохранение DataFrame с результатами в Excel
sales_forecast.to_excel('result/Прогнозы_Holt-Winters_MulMul.xlsx')