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

# Создание словаря для хранения слагов и их статистики
slug_stats = {}

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
        if slug not in slug_stats:
            slug_stats[slug] = {'count': 1, 'visibility_sum': visibility, 'nesting_level': i + 1, 'example_url': url}
        else:
            slug_stats[slug]['count'] += 1
            slug_stats[slug]['visibility_sum'] += visibility

# Создание DataFrame из словаря
result_df = pd.DataFrame.from_dict(slug_stats, orient='index')
result_df.reset_index(inplace=True)
result_df.columns = ['Слаг', 'Количество', 'Сум.видимость', 'Глубина', 'Пример URL']

# Сортировка по колонке "Количество" (count)
result_df = result_df.sort_values(by='Количество', ascending=False)

# Сохранение результата в Excel с заданными ширинами колонок
export_file = 'export.xlsx'
with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
    result_df.to_excel(writer, index=False, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']
    
    # Задаем ширину колонок
    column_widths = [20, 20, 20, 20, 30]  # Ширина колонок
    for i, column in enumerate(worksheet.columns):
        max_length = max(len(str(cell.value)) for cell in column)
        worksheet.column_dimensions[column[0].column_letter].width = column_widths[i]

# Вычисляем статистику
total_urls_processed = df.shape[0]
unique_slugs = result_df.shape[0]
elapsed_time = time.time() - start_time

# Выводим статистику
print(f'Обработано {total_urls_processed} URL')
print(f'Уникальных слагов: {unique_slugs}')
print(f'Время выполнения скрипта: {elapsed_time} секунд')

# Ожидание нажатия Enter
input('Нажмите Enter для завершения...')

# Открываем файл после выполнения скрипта
os.system(f'start excel.exe "{export_file}"')
