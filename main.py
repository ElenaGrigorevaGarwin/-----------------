import os
import pandas as pd

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