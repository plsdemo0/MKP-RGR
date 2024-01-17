import array

import openpyxl
import numpy as np

file_path = 'МКП_РГР.xlsx'
workbook = openpyxl.load_workbook(file_path)

# Выбираем лист "Data" или создаем его, если его нет
sheet_name = 'Data'
if sheet_name in workbook.sheetnames:
    sheet = workbook[sheet_name]
    sheet.delete_rows(1, sheet.max_row)
else:
    sheet = workbook.create_sheet(title=sheet_name)

text_file_path = 'Данные'
with open(text_file_path, 'r') as file:
    # Читаем строки из файла
    lines = file.readlines()
    # Разбиваем строки на части (предполагаем, что разделены пробелами)
    for line_number, line in enumerate(lines, start=1):
        data_parts = line.split()
        # Записываем данные в соответствующие ячейки в лист "Data"
        for col_number, data_part in enumerate(data_parts, start=1):
            cell = sheet.cell(row=line_number, column=col_number)
            cell.value = data_part
            # Пробуем преобразовать данные в числовой формат
            try:
                cell.value = float(
                    data_part.replace(',', '.'))  # Заменяем запятую на точку для правильного преобразования
            except ValueError:
                # Если не удалось преобразовать в число, оставляем строку
                cell.value = data_part

# Сохраняем изменения в файле
workbook.save(file_path)
print(f"Данные успешно записаны в лист {sheet_name} файла {file_path}")

# Создание массива для записи данных КА для подсчетов пересечений

# Запись названий спутников
data_list = []
for i in range(1,29,3):
    name_list = []
    cell = sheet.cell(row = i, column=1)
    name_list.append(cell.value)
    data_list.append(name_list)
# Запись данных для каждого спутника
for j in range(len(data_list)):
    for d in range(2,9):
        cell = sheet.cell(row = (j+1)*3, column=d)
        data_list[j].append(cell.value)

# Константы
nu = 396628
R = 6378.16
J2 = 0.0010827
eps = 3/2 * nu * J2 * R**2

# Рассчитываем Омеги, сначала Град/Виток, затем Град/Сутки
for i in range(len(data_list)):
    omega_1 = -(2*np.pi*eps) / nu / (((nu * (1 / (data_list[i][7] / 86400))**2/4/np.pi**2)**(1/3)) * (1-data_list[i][4]**2))**2 * np.cos(data_list[i][2]*np.pi/180)
    data_list[i].append(omega_1)
    omega_2 = data_list[i][-1]/(1 / (data_list[i][7] / 86400)) * 86400
    data_list[i].append(omega_2)
# Вывод массива данных КА
for i in range(len(data_list)):
    print(data_list[i])
# Закрытие книги после использования
workbook.close()

