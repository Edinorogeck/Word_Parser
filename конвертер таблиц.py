import pandas as pd
import docx

# Открываем файл
doc = docx.Document(input('Введите путь к файлу: '))

# Создаем пустой список для хранения таблиц
tables = []

# Итерируемся по всем элементам документа
for element in doc.element.body:

    # Если элемент является таблицей
    if isinstance(element, docx.table.Table):
        # Извлекаем данные из таблицы и преобразуем их в DataFrame
        data = [[cell.text for cell in row.cells] for row in element.rows]
        df = pd.DataFrame(data[1:], columns=data[0])
        df.to_csv('convert_output.csv', encoding='windows-1251', index=False, sep=';')
        # Добавляем DataFrame в список таблиц
        tables.append(df)
