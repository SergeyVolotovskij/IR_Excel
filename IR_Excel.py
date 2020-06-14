#импортируем необходимые библиотеки
from openpyxl import workbook
from openpyxl.styles import Font, Color, colors
from openpyxl import load_workbook
import openpyxl

#для удобства вводим название файла
# filename_0 = input("ВВЕДИТЕ ИМЯ ФАЙЛА: ")
filename = "Список.xlsx"

#вытянули данные с документа
active_excel = load_workbook(filename=filename,data_only=True)#data_only=True

#делаем или смотрим активный лист
active_sheet = active_excel.active

#нужно понять максимальный размер данных на листе
max_row = active_sheet.max_row
max_column = active_sheet.max_column

print("КОЛИЧЕСТВО СТРОК: " + str(max_row))
print("КОЛИЧЕСТВО КОЛОНОК: " + str(max_column))

# запишем наименование 7, 8 колонки
active_sheet["G1"] = 'Пробел'
active_sheet["H1"] = 'Сцепить'

# вносим данные
space = " "
# combine = active_sheet['B' + str(i)]

#делаем цикл по заполнению пробелом всего диапазона колонки
for i in range(2,(max_row)):
    _= active_sheet.cell(column= max_column - 1, row=i, value=space)

#делаем цикл по заполнению СЦЕПИТЬ всего диапазона колонки в екселе
for i in range(2,(max_row)):
    a = 'A' + str(i)
    b = 'B' + str(i)
    c = 'G' + str(i)

    d = '=' + a + '&' + c + '&' + b
    _= active_sheet.cell(column= max_column, row=i, value=d)

#делаем цикл по заполнению СЦЕПИТЬ всего диапазона колонки в нашем списке
spisok = []
for i in range(2,(max_row)):
    a_s = active_sheet["A" + str(i)].value
    b_s = active_sheet["B" + str(i)].value
    g_s = active_sheet["G" + str(i)].value

    d_s = str(a_s) + str(g_s) + str(b_s)
    spisok.append(d_s)
    # print(d)

#анализируем ШК
barcode = []
for i in range(2,(max_row)):
    e = active_sheet["E" + str(i)].value
    barcode.append(e)

unique_barcode = []
for i in barcode:
    if barcode.count(i) == 1: #если количество вхождений = 1 - элемент уникальный
        unique_barcode.append(i)

if len(barcode) == len(unique_barcode):
    print("ДУБЛИКАТЫ ШК ОТСУТСТВУЮТ!")
else:
    print("ЕСТЬ ДУБЛИ ШК!")

#анализируем наименование
unique_spisok = []
for i in spisok:
    if spisok.count(i) == 1:
        unique_spisok.append(i)

if len(spisok) == len(unique_spisok):
    print("ДУБЛИКАТЫ НАИМЕНОВАНИЙ ОТСУТСТВУЮТ!")
else:
    print("ЕСТЬ ДУБЛИ НАИМЕНОВАНИЙ!")

# сохраняем изменения
active_excel.save("Список.xlsx") #сохраняем все изменения
