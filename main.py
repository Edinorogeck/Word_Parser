
import docx
import pandas as pd
import re

# Открываем файл и ищем все таблицы
document = docx.Document(input('Введите путь к файлу: '))
# Инициализируем список для хранения всех таблиц

tables = document.tables

print(len(tables))


#print(len(tables))


data = []
# Обходим все таблицы в документе и выводим на экран их содержимое
for table in tables:
    for row in table.rows:
        row_data = []
        for cell in row.cells:
            text = cell.text
            row_data.append(text)
            if 'МКИО' in text:
                print(text)
        data.append(row_data)

df = pd.DataFrame(data)

df.to_csv('output.csv', encoding='windows-1251', index=False, sep=';')


