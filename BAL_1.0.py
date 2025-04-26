# -*- coding: utf-8 -*-
import openpyxl
import os
import subprocess

# Получаем текущую директорию
current_directory = os.path.dirname(os.path.abspath(__file__))

# Находим все файлы Excel в текущей директории
excel_files = [f for f in os.listdir(current_directory) if f.endswith('.xlsx')]

# Проверяем, что файл Excel существует и единственный
if len(excel_files) != 1:
    raise FileNotFoundError("В текущей директории должен быть только один файл Excel.")

# Полный путь к файлу Excel
excel_file_path = os.path.join(current_directory, excel_files[0])

# Открываем файл Excel
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active

# Новые значения, которые нужно вставить в столбец 3
new_values = {
    '1СТ': 'Бевза Алексей Леонидович',  # Вставьте новое значение для кода 1СТ
    '2СТ': 'образование высшее медицинское, специальность – судебно-медицинская экспертиза,',  # Вставьте новое значение для кода 2СТ
    '3СТ': 'стаж работы по специальности с 1.09.2004 г., высшая квалификационная категория,',  # Вставьте новое значение для кода 3СТ
    '4СТ': 'кандидат медицинских наук, врач судебно-медицинский эксперт',  # Вставьте новое значение для кода 4СТ
    '5СТ': ' ',  # Вставьте новое значение для кода 5СТ
    2: 'F:\\Работа\\Автоматизация-5\\2',  # Вставьте новый путь для кода 2
    3: 'F:\\Работа\\Автоматизация-5\\3',  # Вставьте новый путь для кода 3
    4: 'F:\\Работа\\Автоматизация-5\\4',   # Вставьте новый путь для кода 4
    5: 'F:\\Работа\\Автоматизация-5\\5',  # Вставьте новый путь для кода 5
    6: 'F:\\Работа\\Автоматизация-5\\6',   # Вставьте новый путь для кода 6
}

# Проходим по строкам и обновляем значения в столбце 3 для заданных кодов
for row in sheet.iter_rows(min_row=2, max_col=1):  # Начинаем с второй строки (первая - заголовок)
    code = row[0].value
    if code in new_values:
        sheet.cell(row=row[0].row, column=3).value = new_values[code]

# Сохраняем изменения в файле Excel
workbook.save(excel_file_path)

# Указываем путь к запускаемой программе
program_to_run = os.path.normpath('F:\\Работа\\Автоматизация-5\\cmd\\V1.10.py')

# Проверяем наличие файла
if not os.path.exists(program_to_run):
    raise FileNotFoundError(f"Файл {program_to_run} не найден.")

# Запускаем другую программу
subprocess.run(['python', program_to_run])
