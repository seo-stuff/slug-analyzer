# Slug Analyzer (ft Keys.so)

## Описание
Задача скрипта - проводить анализ отчета о страницах из сервиса Keys.so для выявления сильных сторон конкурентов на основе структуры URL и их видимости.
  По промокоду **heymm20** можно получить скидку в 20% на первую оплату для новых пользователей.

![Результаты](/images/demo.png)
![Результаты](/images/demo2.png)

## Начало работы
* Установите Python с официального сайта - https://www.python.org/downloads/windows/
* Запустите установщик Python и следуйте инструкциям по установке. Убедитесь, что установщик установил Python в системный путь (добавил Python в переменную PATH).
* Через командную строку в Windows (CMD) установить библиотеки:
> pip install pandas chardet openpyxl
* Скачайте отчет по страницам для любого сайта из сервиса Keys.so (кодировка любая)
* Сохраните файл с именем import.csv и положите рядом со скриптом
* Запустите файл скрипта и дождитесь выполнения

## История изменений
### Загружена версия 2.0 (22.01.2024)
* Присутствует общий лист с данными
* Создаются отдельные листы для каждого уровня вложенности для удобного изучения
* Вывод статистики после выполнения скрипта (время работы, кол-во обработанных URL)
* Открытие файла после завершения

### Загружена версия 1.0 (22.01.2024)
Базовая версия скрипта со следующими данными на одном листе:
* Слаг
* Количество
* Суммарная видимость всех URL со слагом
* Глубина вложенности
* Пример URL (для ознакомления)
