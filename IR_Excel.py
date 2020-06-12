#импортируем необходимые библиотеки
from openpyxl import workbook
from openpyxl.styles import Font, Color, colors
from openpyxl import load_workbook

#для удобства вводим название файла
# filename_0 = input("ВВЕДИТЕ ИМЯ ФАЙЛА: ")
filename = "Список.xlsx"
# print(filename)

#вытянули данные с документа
active_excel = load_workbook(filename=filename, data_only=True)

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




# сохраняем изменения
active_excel.save("Список.xlsx") #сохраняем все изменения
print('ИЗМЕНЕНИЯ СОХРАНЕНЫ')