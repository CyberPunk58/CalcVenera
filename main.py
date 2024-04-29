import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import re

def process_excel():
    global file_path
    if not file_path:
        messagebox.showerror("Ошибка", "Пожалуйста, выберите файл Excel.")
        return

    try:
        workbook = openpyxl.load_workbook(file_path)
        main_sheet = workbook.active

        # Создание или очистка листа "Специалисты"
        sheet_name = 'Специалисты'
        if sheet_name in workbook.sheetnames:
            new_sheet = workbook[sheet_name]
            new_sheet.delete_rows(1, new_sheet.max_row)
        else:
            new_sheet = workbook.create_sheet(sheet_name)

        # Обработка данных
        delimiter_pattern = r',|\.|:'
        unique_entries = set()

        for cell in main_sheet['C']:
            items = re.split(delimiter_pattern, cell.value.strip()) if cell.value else []
            for item in items:
                cleaned = item.strip()
                if cleaned:
                    unique_entries.add(cleaned)

        for i, entry in enumerate(sorted(unique_entries), start=1):
            new_sheet.cell(row=i, column=1).value = entry

        # Специалисты и фразы
        special_phrases = ["гинеколог", "Мазок на флору", "Мазок на онкоцитологию", "УЗИ органов малого таза"]
        specialists_sheet = new_sheet

        for i in range(1, specialists_sheet.max_row + 1):
            specialist = specialists_sheet.cell(row=i, column=1).value
            sum_values = 0

            # Проверка и подсчет
            for row in main_sheet.iter_rows(min_row=1):
                if row[2].value and specialist in str(row[2].value):
                    check_special = any(phrase in specialist for phrase in special_phrases)
                    if check_special:
                        value2 = row[5].value if row[5].value is not None and isinstance(row[5].value, (int, float)) else 0
                        sum_values += value2
                    else:
                        value1 = row[4].value if row[4].value is not None and isinstance(row[4].value, (int, float)) else 0
                        value2 = row[5].value if row[5].value is not None and isinstance(row[5].value, (int, float)) else 0
                        sum_values += value1 + value2

            specialists_sheet.cell(row=i, column=2).value = sum_values

        # Сохранение файла с результатом
        workbook.save(file_path.replace('.xlsx', '_calculation.xlsx'))
        messagebox.showinfo("Результат", "Файл успешно обработан!")

    except Exception as e:
        messagebox.showerror("Ошибка при обработке", str(e))

def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_label.config(text=f"Выбранный файл: {file_path}")

# GUI
root = tk.Tk()
root.title("Обработка Excel")
root.geometry("500x200")

file_path = ''

# Элементы интерфейса
select_button = tk.Button(root, text="Выбрать файл Excel", command=select_file)
select_button.pack(pady=10)

file_label = tk.Label(root, text="Файл не выбран")
file_label.pack()

process_button = tk.Button(root, text="Начать обработку", command=process_excel)
process_button.pack(pady=20)

root.mainloop()
