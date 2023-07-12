import docx
import pandas as pd


# GLOBALS PARAMETRS

# Массив для правильных DataFrame pandas
df_MKIO = []
df_ETH = []
# Массив названий МКИО сообщений
arrayMessageNames_MKIO = []
arrayMessageNames_ETH = []


# Функция проверки на вхождение подстроки "МКИО."
def Check_MKIO(text):
    words = text.split()
    for word in words:
        if "МКИО." in word:
            return True

# Функция проверки на вхождение подстроки "Eth."
def Check_ETH(text):
    words = text.split()
    for word in words:
        if "Eth." in word:
            return True

# Функция возвращает название сообщения
def GetMessageName(text):
    words = text.split()
    for word in words:
        if "Eth." in word:
            return word
        if "МКИО." in word:
            return word

# Функция создания списка имен сообщений МКИО
def GetArrayMessageNames_MKIO(df_MKIO):
    # Алгоритм создания списка имен сообщений МКИО
    for i, table_MKIO in enumerate(df_MKIO):
        messageName = ''
        # Цикл для строк
        for index, row in table_MKIO.iterrows():
            # Цикл для колонок
            for col in table_MKIO.columns:
                text = table_MKIO.loc[index, col]
                if Check_MKIO(text):
                    messageName = GetMessageName(text)
                    break
                if Check_ETH(text):
                    messageName = GetMessageName(text)
                    break
            if messageName != '':
                arrayMessageNames_MKIO.append(messageName)
                break

# Функция создания списка имен сообщений Ethernet
def GetArrayMessageNames_ETH(df_ETH):
    # Алгоритм создания списка имен сообщений Ethernet
    for i, table_ETH in enumerate(df_ETH):
        messageName = ''
        # Цикл для строк
        for index, row in table_ETH.iterrows():
            # Цикл для колонок
            for col in table_ETH.columns:
                text = table_ETH.loc[index, col]
                if Check_MKIO(text):
                    messageName = GetMessageName(text)
                    break
                if Check_ETH(text):
                    messageName = GetMessageName(text)
                    break
            if messageName != '':
                arrayMessageNames_MKIO.append(messageName)
                break




# Открываем файл формата .docx
doc = docx.Document(input('Введите путь к файлу: '))


# Массив для DataFrame pandas
df = []

# Извлекаем все таблицы из документа и сохраняем их в CSV-файлы
for i, table in enumerate(doc.tables):
    # Массив для хранения массивов(по итогу двумерный массив)
    data_row = []
    # Создаем DataFrame из данных таблицы
    for row in table.rows:
        data_cell = []
        for cell in row.cells:
            data_cell.append(cell.text)
        data_row.append(data_cell)
    # Добавляем таблицу(двуменрный массив) в массив для DataFrame pandas
    df.append(pd.DataFrame(data_row))


# Флаги для МКИО и Eth
flag_MKIO = False
flag_ETH = False
# Цикл по всем DataFrame pandas
for i, table in enumerate(df):
    # Цикл для строк
    for index, row in table.iterrows():
        # Цикл для колонок
        for col in table.columns:
            '''
            В общем и целом ищется первое вхождение кючевого слова
            для МКИО сообщения "МКИО."
            для Ethernet сообщения "Eth."
            '''
            text = table.loc[index, col]
            if Check_MKIO(text):
                flag_MKIO = True
                break
            if Check_ETH(text):
                flag_ETH = True
                break
        if flag_MKIO == True:
            df_MKIO.append(table)
            flag_MKIO = False
            break
        if flag_ETH == True:
            df_ETH.append(table)
            flag_MKIO = False
            break



print("---------------------------------------------------------")

GetArrayMessageNames_MKIO(df_MKIO)
GetArrayMessageNames_ETH(df_ETH)




print(arrayMessageNames_MKIO)


# Сохраняем DataFrame в CSV-файл
for i, table in enumerate(df_MKIO):
    table.to_csv(f'new_table_{i+1}.csv', encoding='windows-1251', index=False, sep=';')
