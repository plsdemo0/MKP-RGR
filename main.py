import openpyxl

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