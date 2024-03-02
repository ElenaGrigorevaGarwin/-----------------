# Преобразование файла с продажами для моделей прогноза 
import os
import pandas as pd
import importlib.util

df = pd.read_excel('Факт отгрузок.xlsx')


df_transposed = df.T  # Транспонирование таблицы
df_transposed.columns = df_transposed.iloc[0]  # Сделать первую строку заголовком
df_transposed = df_transposed.iloc[1:]  # Удалить первую строку после использования как заголовка

# Сброс индекса и перевод его в обычный столбец, а также присвоение имени столбцу "Дата"
df_transposed.reset_index(inplace=True)
df_transposed.rename(columns={'index': 'Дата'}, inplace=True)

# Заполнение отсутствующих данных нулями
df_transposed.fillna(0, inplace=True)

df_transposed.to_excel('Факт отгрузок.xlsx', index=False)  # запись итогового файла без индекса

# Запускаем прогнозы
# Получаем список файлов в папке "model"
model_files = os.listdir("model")

# Перебираем каждый файл
for file_name in model_files:
    # Проверяем, что это файл с расширением .py
    if file_name.endswith(".py"):
        # Формируем путь к файлу
        file_path = os.path.join("model", file_name)
        print(f"Запуск модели из файла: {file_path}")
        
        # Загружаем модуль
        module_name = file_name[:-3]  # Убираем расширение .py
        spec = importlib.util.spec_from_file_location(module_name, file_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        print("Модель успешно выполнена.")
    
# Преобразование всех прогнозов для лучшего метода

def transpose_tables_in_folder(folder_path):
    
    files = os.listdir(folder_path)  # Получить список файлов в указанной папке
    
    # Итерирование по каждому файлу Excel
    for file in files:
        file_path = os.path.join(folder_path, file) # Чтение файлов
        print(f"Транспонирую файл: {file_path}")
        df = pd.read_excel(file_path)
        df_transposed = df.T # Транспонирование таблицы
        df_transposed.columns = df_transposed.iloc[0] # Сделать первую строку заголовком
        df_transposed = df_transposed.iloc[1:]  # Удалить первую строку после использования как заголовка
        df_transposed.index.name = "Отдел" # Изменить наименование первой колонки
        df_transposed.to_excel(file_path, index=True) # запись итогового файла вместо исходного
        print(f"Таблица транспонирована и сохранена обратно в: {file_path}")

# Применение функции к папке "result"
folder_path = "result"
transpose_tables_in_folder(folder_path)