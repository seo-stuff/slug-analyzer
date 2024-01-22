import pandas as pd
import chardet
from urllib.parse import urlparse
import time
import os
from openpyxl import load_workbook

# Определение кодировки файла
def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

# Получение кодировки файла
file_encoding = detect_encoding('import.csv')

# Чтение данных из CSV-файла с определенной кодировкой
df = pd.read_csv('import.csv', sep=';', encoding=file_encoding)

# Удаление лишних символов из имен столбцов
df.columns = df.columns.str.strip()

# Создание словаря для хранения данных по уровням вложенности
nested_slug_stats = {}

# Функция для извлечения всех слагов после первого слеша после начала домена
def extract_slugs(url):
    parsed_url = urlparse(url)
    path = parsed_url.path
    parts = path.split('/')
    slugs = [parts[i] for i in range(1, len(parts)) if parts[i] != '']
    return slugs

# Засекаем время начала выполнения скрипта
start_time = time.time()

# Обработка каждой строки и подсчет статистики
for index, row in df.iterrows():
    url = row['url']
    slugs = extract_slugs(url)
    visibility = int(row['Ср. видимость'])
    
    for i, slug in enumerate(slugs):
        if i + 1 not in nested_slug_stats:
            nested_slug_stats[i + 1] = {'Слаг': [], 'Количество': [], 'Сум.видимость': [], 'Глубина': [], 'Пример URL': []}
        
        nested_slug_stats[i + 1]['Слаг'].append(slug)
        nested_slug_stats[i + 1]['Количество'].append(1)
        nested_slug_stats[i + 1]['Сум.видимость'].append(visibility)
        nested_slug_stats[i + 1]['Глубина'].append(i + 1)
        nested_slug_stats[i + 1]['Пример URL'].append(url)

# Создаем Excel-файл с отдельными листами для каждого уровня вложенности
export_file = 'export.xlsx'
with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
    # Создаем первый лист со всеми данными
    df_all = pd.concat([pd.DataFrame(data) for data in nested_slug_stats.values()], ignore_index=True)
    df_all = df_all.groupby('Слаг').agg({'Количество': 'sum', 'Сум.видимость': 'sum', 'Глубина': 'first', 'Пример URL': 'first'}).reset_index()
    df_all = df_all.sort_values(by='Сум.видимость', ascending=False)
    df_all['% от общего'] = round((df_all['Сум.видимость'] / df_all['Сум.видимость'].sum()) * 100, 1)  # Округляем проценты
    df_all = df_all[['Слаг', 'Сум.видимость', 'Количество', 'Глубина', 'Пример URL', '% от общего']]  # Переупорядочиваем колонки
    df_all.to_excel(writer, index=False, sheet_name='Статистика')
    worksheet = writer.sheets['Статистика']
    
    # Задаем ширину колонок
    column_widths = [20, 20, 20, 20, 30, 20]  # Ширина колонок
    for i, column in enumerate(worksheet.columns):
        max_length = max(len(str(cell.value)) for cell in column)
        worksheet.column_dimensions[column[0].column_letter].width = column_widths[i]

    # Создаем отдельные листы для каждого уровня вложенности
    for nesting_level, data in nested_slug_stats.items():
        df_nesting_level = pd.DataFrame(data)
        df_nesting_level = df_nesting_level.groupby('Слаг').agg({'Количество': 'sum', 'Сум.видимость': 'sum', 'Глубина': 'first', 'Пример URL': 'first'}).reset_index()
        df_nesting_level = df_nesting_level.sort_values(by='Сум.видимость', ascending=False)
        df_nesting_level['% от общего'] = round((df_nesting_level['Сум.видимость'] / df_all['Сум.видимость'].sum()) * 100, 1)
        df_nesting_level = df_nesting_level[['Слаг', 'Сум.видимость', 'Количество', 'Глубина', 'Пример URL', '% от общего']]
        df_nesting_level.to_excel(writer, index=False, sheet_name=f'Вложенность {nesting_level}')
        worksheet = writer.sheets[f'Вложенность {nesting_level}']

        # Задаем ширину колонок
        for i, column in enumerate(worksheet.columns):
            max_length = max(len(str(cell.value)) for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = column_widths[i]

# Вычисляем статистику
total_urls_processed = df.shape[0]
elapsed_time = time.time() - start_time

# Выводим статистику
print(f'Обработано {total_urls_processed} URL')
print(f'Время выполнения скрипта: {elapsed_time} секунд')

# Ожидание нажатия Enter
input('Нажмите Enter для завершения...')

# Открываем файл после выполнения скрипта
os.system(f'start excel.exe "{export_file}"')
