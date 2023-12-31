import docx
import pandas as pd
import os


# GLOBALS PARAMETRS

# Массив для рабочих DataFrame pandas
df_MKIO = []
df_ETH = []
# Массив для итоговых DataFrame pandas
df_Total_MKIO = []
df_Total_ETH = []
# Массив названий МКИО и ETH сообщений
arrayMessageNames_MKIO = []
arrayMessageNames_ETH = []
# Массив адресов МКИО и ETH сообщений
arrayMessageAddress_MKIO = []
arrayMessageAddress_ETH = []
# Массив с полезными данными МКИО и ETH сообщений
arrayMessageData_MKIO = []
arrayMessageData_ETH = []


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


# Функция создания списка имен сообщений ETH
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
                arrayMessageNames_ETH.append(messageName)
                break


# Функция создания списка адресов сообщений МКИО
def GetArrayMessageAddress_MKIO(df_MKIO):
    # Алгоритм создания списка имен сообщений МКИО
    for i, table_MKIO in enumerate(df_MKIO):
        arrayMessageAddress_MKIO.append(table_MKIO.loc[1, 2].split(None, 1)[0])


# Функция создания списка адресов сообщений ETH
def GetArrayMessageAddress_ETH(df_ETH):
    # Алгоритм создания списка имен сообщений МКИО
    for i, table_ETH in enumerate(df_ETH):
        addressData = []
        text = table_ETH.loc[1, 2].replace("\n", " ")
        addressData.append(text.split()[0])
        addressData.append(text.split()[1])
        arrayMessageAddress_ETH.append(addressData)


# Функция создания массива с полезными данными ETH сообщений
def GetArrayMessageData_MKIO(df_MKIO):
    for i, table_MKIO in enumerate(df_MKIO):
        data_rows = []
        for row in range(4, len(table_MKIO)):
            data_cells = []
            for cell in range(1, 3):
                data_cells.append(table_MKIO.loc[row, cell])
            data_rows.append(data_cells)
        arrayMessageData_MKIO.append(data_rows)


# Функция создания массива с полезными данными ETH сообщений
def GetArrayMessageData_ETH(df_ETH):
    for i, table_ETH in enumerate(df_ETH):
        data_rows = []
        for row in range(4, len(table_ETH)):
            data_cells = []
            for cell in range(1, 3):
                data_cells.append(table_ETH.loc[row, cell])
            data_rows.append(data_cells)
        arrayMessageData_ETH.append(data_rows)


# Функция создания файлов MKIO сообщений
def MakeMessageCSVFiles_MKIO(arrayMessageNames_MKIO, arrayMessageAddress_MKIO, arrayMessageData_MKIO, folder_name):
    data_df = pd.DataFrame()
    for i in range(len(arrayMessageNames_MKIO)):
        name = arrayMessageNames_MKIO[i]
        data_df.at[0, 0] = arrayMessageNames_MKIO[i]
        data_df.at[1, 0] = arrayMessageAddress_MKIO[i]
        for i, data in enumerate(arrayMessageData_MKIO[i]):
            if data[0] == "":
                continue
            if data[1] == "":
                continue
            data_df.at[i + 2, 0] = data[0]
            data_df.at[i + 2, 1] = data[1]
        data_df.to_csv(f'{folder_name}/{name}.csv', encoding='windows-1251', index=False, sep=';')


# Функция создания файлов ETH сообщений
def MakeMessageCSVFiles_ETH(arrayMessageNames_ETH, arrayMessageAddress_ETH, arrayMessageData_ETH, folder_name):
    data_df = pd.DataFrame()
    for i in range(len(arrayMessageNames_ETH)):
        name = arrayMessageNames_ETH[i]
        data_df.at[0, 0] = arrayMessageNames_ETH[i]
        mesageAddress = arrayMessageAddress_ETH[i]
        data_df.at[1, 0] = mesageAddress[0]
        data_df.at[1, 1] = mesageAddress[1]
        for i, data in enumerate(arrayMessageData_ETH[i]):
            if data[0] == "":
                continue
            if data[1] == "":
                continue
            data_df.at[i + 2, 0] = data[0]
            data_df.at[i + 2, 1] = data[1]
        data_df.to_csv(f'{folder_name}/{name}.csv', encoding='windows-1251', index=False, sep=';')






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
# Цикл по всем DataFrame pandas для создания массивов таблиц МКИО и ETH
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
            flag_ETH = False
            break




folder_name = 'Configuration folder'
if not os.path.exists(folder_name):
    os.mkdir(folder_name)








GetArrayMessageNames_MKIO(df_MKIO)
GetArrayMessageNames_ETH(df_ETH)

GetArrayMessageAddress_MKIO(df_MKIO)
GetArrayMessageAddress_ETH(df_ETH)

GetArrayMessageData_MKIO(df_MKIO)
GetArrayMessageData_ETH(df_ETH)

MakeMessageCSVFiles_MKIO(arrayMessageNames_MKIO, arrayMessageAddress_MKIO, arrayMessageData_MKIO, folder_name)
MakeMessageCSVFiles_ETH(arrayMessageNames_ETH, arrayMessageAddress_ETH, arrayMessageData_ETH, folder_name)



file_path = f'{folder_name}/Configuration_file.txt'

with open(file_path, 'w') as file:
    for name in arrayMessageNames_MKIO:
        file.write(f'{name}.csv\n')
    for name in arrayMessageNames_ETH:
        file.write(f'{name}.csv\n')




'''

# Сохраняем DataFrame в CSV-файл
for i, table in enumerate(df_MKIO):
    table.to_csv(f'MKIO_table_{i+1}.csv', encoding='windows-1251', index=False, sep=';')

for i, table in enumerate(df_ETH):
    table.to_csv(f'ETH_table_{i+1}.csv', encoding='windows-1251', index=False, sep=';')

'''
