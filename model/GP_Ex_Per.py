# модель с экспоненциальным квадратом и периодическим ядрами
import GPy
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

df = pd.read_excel('Факт отгрузок.xlsx')
df['Дата'] = pd.to_datetime(df['Дата'])
df.set_index('Дата', inplace=True)

predictions = pd.DataFrame(index=pd.date_range(start=df.index[-1] + pd.DateOffset(months=1), periods=12, freq='MS'))

for department in df.columns:
  series = df[department].values.reshape(-1, 1)
  X = np.arange(len(series)).reshape(-1, 1) # Создание временной шкалы
  
  # Создание и обучение модели Гауссовского процесса с экспоненциальным квадратом и периодическим ядрами
  kernel_exp = GPy.kern.Exponential(input_dim=1)
  kernel_periodic = GPy.kern.StdPeriodic(input_dim=1)
  kernel = kernel_exp + kernel_periodic  # Комбинация ядер
  model = GPy.models.GPRegression(X, series, kernel)

  model.Gaussian_noise.variance = 0.1  # Устанавливаем дисперсию шума
  model.optimize(messages=True) # Оптимизация параметров модели

  # Предсказание на всей временной шкале
  X_pred = np.arange(len(series) + 12).reshape(-1, 1)  # Предсказание на 12 месяцев вперед
  y_pred, sigma = model.predict(X_pred)

  predictions[department] = y_pred[-12:]  # Берем только последние 12 прогнозов


predictions.to_excel("result/Прогнозы_GP_Ex_Per.xlsx")