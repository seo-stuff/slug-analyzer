import pandas as pd
from urllib.parse import urlparse
import time
import os
import pyfiglet
from openpyxl import load_workbook
from datetime import datetime  # Добавляем импорт модуля datetime

# Проверка наличия файла import.csv
import_file = 'import.csv'

if not os.path.isfile(import_file):
    print(f'!!!Файл {import_file} не найден!!!')
    print(f'Скачайте отчет по страницам для сайта из сервиса Keys.so и поместите его в текущую директорию.')
    input('\nНажмите Enter для завершения...')
    exit()  # Выход из скрипта

# Запуск скрипта
text = "Slug Analyzer ver. 2.2"
ascii_art = pyfiglet.figlet_format(text, font="slant")
print(ascii_art)
print(f'Скрипт запущен, ожидайте ...')

# Определение кодировки файла
def detect_encoding(file_path):
    encodings_to_try = ['utf-8', 'cp1251']
    detected_encoding = None
    
    for encoding in encodings_to_try:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read()
            detected_encoding = encoding
            break
        except UnicodeDecodeError:
            continue
    
    if detected_encoding is None:
        raise Exception('Не удалось определить кодировку файла')
    
    return detected_encoding

# Получение кодировки файла
file_encoding = detect_encoding(import_file)

# Чтение данных из CSV-файла с определенной кодировкой
df = pd.read_csv(import_file, sep=';', encoding=file_encoding)

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
total_rows = len(df)
for index, row in df.iterrows():
    url = row['url']
    slugs = extract_slugs(url)
    visibility = int(row['Ср. видимость'])
    
    for i, slug in enumerate(slugs):
        if i + 1 not in nested_slug_stats:
            nested_slug_stats[i + 1] = {'Слаг': [], 'Сум.видимость': [], 'Количество': [], 'Глубина': [], 'Пример URL': []}
        
        nested_slug_stats[i + 1]['Слаг'].append(slug)
        nested_slug_stats[i + 1]['Сум.видимость'].append(visibility)
        nested_slug_stats[i + 1]['Количество'].append(1)
        nested_slug_stats[i + 1]['Глубина'].append(i + 1)
        nested_slug_stats[i + 1]['Пример URL'].append(url)

    # Вычисляем процент выполнения и оставшееся время
    progress = (index + 1) / total_rows * 100
    elapsed_time = time.time() - start_time
    estimated_total_time = (elapsed_time / (index + 1)) * total_rows - elapsed_time

    # Форматируем и выводим статус выполнения
    status_str = (
        f"[{progress:.2f}%] Обработано {index + 1}/{total_rows} строк | "
        f"Оставшееся время: {estimated_total_time:.2f} сек."
    )
    print(status_str, end='\r')

# Генерируем строку с текущей датой и временем в читаемом формате
current_datetime = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')

# Создаем имя файла с датой и временем
export_file = f'export_{current_datetime}.xlsx'

# Создаем Excel-файл с отдельными листами для каждого уровня вложенности
print(f'\nИдёт запись в файл {export_file} ...')
with pd.ExcelWriter(export_file, engine='openpyxl') as writer:
    # Создаем первый лист со всеми данными
    df_all = pd.concat([pd.DataFrame(data) for data in nested_slug_stats.values()], ignore_index=True)
    df_all = df_all.groupby('Слаг').agg({'Сум.видимость': 'sum', 'Количество': 'sum', 'Глубина': 'first', 'Пример URL': 'first'}).reset_index()
    df_all['Ценность'] = round(df_all['Сум.видимость'] / df_all['Количество'])  # Округление до целых чисел
    df_all = df_all.sort_values(by='Сум.видимость', ascending=False)
    df_all.to_excel(writer, index=False, sheet_name='Статистика')
    worksheet = writer.sheets['Статистика']

    # Задаем ширину колонок
    column_widths = [20, 20, 20, 20, 20, 30]  # Ширина колонок
    for i, column in enumerate(worksheet.columns):
        max_length = max(len(str(cell.value)) for cell in column)
        worksheet.column_dimensions[column[0].column_letter].width = column_widths[i]

    # Создаем отдельные листы для каждого уровня вложенности
    for nesting_level, data in nested_slug_stats.items():
        df_nesting_level = pd.DataFrame(data)
        df_nesting_level = df_nesting_level.groupby('Слаг').agg({'Сум.видимость': 'sum', 'Количество': 'sum', 'Глубина': 'first', 'Пример URL': 'first'}).reset_index()
        df_nesting_level['Ценность'] = round(df_nesting_level['Сум.видимость'] / df_nesting_level['Количество'])  # Округление до целых чисел
        df_nesting_level = df_nesting_level.sort_values(by='Сум.видимость', ascending=False)
        df_nesting_level.to_excel(writer, index=False, sheet_name=f'Вложенность {nesting_level}')
        worksheet = writer.sheets[f'Вложенность {nesting_level}']

        # Задаем ширину колонок
        for i, column in enumerate(worksheet.columns):
            max_length = max(len(str(cell.value)) for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = column_widths[i]

# Вычисляем статистику
total_urls_processed = df.shape[0]
elapsed_time = time.time() - start_time

# Округляем секунды до десятых
elapsed_time = round(elapsed_time, 1)

# Выводим статистику
print(f'\nУспешно обработано {total_urls_processed} URL')
print(f'Время выполнения скрипта: {elapsed_time} секунд')

# Ожидание нажатия Enter
input('\nНажмите Enter для открытия файла...')

# Открываем файл после выполнения скрипта
os.system(f'start excel.exe "{export_file}"')
