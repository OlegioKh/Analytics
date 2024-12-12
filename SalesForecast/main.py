import os
import pandas as pd

# Папка з файлами
folder_path = r'C:\Users\kharc\PycharmProjects\Analytics\SalesForecast\files'

# Зчитуємо всі файли Excel у папці
all_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith('.xlsx')]

# Список для об'єднання даних
dataframes = []

# Завантаження кожного файлу і додавання до списку
for file in all_files:
    df = pd.read_excel(file, skiprows=5, names=columns)  # Використовуємо вже визначені назви колонок
    df['Джерело'] = os.path.basename(file)  # Додаємо джерело для відстеження
    dataframes.append(df)

# Об'єднуємо всі файли
final_dataset = pd.concat(dataframes, ignore_index=True)

# Переглядаємо перші рядки
tools.display_dataframe_to_user(name="Об'єднаний датасет продажів", dataframe=final_dataset)