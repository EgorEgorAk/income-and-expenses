import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime

# 🔸 Функция для ввода данных от пользователя
def user_input():
    entries = []  # Список для хранения введённых строк

    # Ввод данных пользователем
    date = input('Введите дату (например, 10.10.2025): ')
    type_ = input('Введите тип (доходы или расходы): ').strip().lower()
    category = input('Введите категорию или описание: ')
    
    # Проверка суммы — должна быть числом
    try:
        amount = int(input("Введите сумму (только число): "))
    except ValueError:
        print("❗ Ошибка: сумма должна быть числом!")
        return []

    # Добавляем введённые данные в список
    entries.append([date, type_, category, amount])
    return entries

# 🔸 Функция для записи в Excel и подсчёта итогов
def income_expenses(data, filename="sample.xlsx"):
    # Если файл уже существует — открываем его
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        # Иначе создаём новый файл
        wb = Workbook()
        ws = wb.active
        ws.title = "Доходы и Расходы"
        
        # Вставим дату и время создания в A1
        ws['A1'] = datetime.now().strftime("Создано: %Y-%m-%d %H:%M:%S")
        
        # Заголовки таблицы
        headers = ['Дата', 'Тип', 'Категория', 'Сумма']
        ws.append([])       # Пустая строка
        ws.append(headers)  # Строка с заголовками

        # Сделаем заголовки жирными
        for cell in ws[3]:
            cell.font = Font(bold=True)

    # Цвета для строк
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # Зелёный — доходы
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")    # Красный — расходы

    # Добавляем новые строки
    for row_data in data:
        ws.append(row_data)  # Добавляем строку
        row_idx = ws.max_row
        fill = green_fill if row_data[1] == "доходы" else red_fill  # Выбор цвета

        # Закрашиваем строку цветом
        for cell in ws[row_idx]:
            cell.fill = fill

    # 🔸 Подсчёт итогов
    total_income = 0
    total_expense = 0

    # Перебираем все строки таблицы, начиная с 4-й
    for row in ws.iter_rows(min_row=4, min_col=2, max_col=4):
        type_cell = row[0].value
        amount_cell = row[2].value

        if isinstance(amount_cell, (int, float)):
            if type_cell == "доходы":
                total_income += amount_cell
            elif type_cell == "расходы":
                total_expense += amount_cell

    balance = total_income - total_expense  # Остаток

        # Удаляем старые итоги (если есть)
    rows_to_delete = []
    for row in ws.iter_rows(min_row=4):
        value = str(row[0].value)
        if value.startswith("Итого") or value.startswith("Остаток"):
            rows_to_delete.append(row[0].row)

    for row_idx in reversed(rows_to_delete):
        ws.delete_rows(row_idx)

    # Удаляем полностью пустые строки
    for row in reversed(range(4, ws.max_row + 1)):
        if all(cell.value is None for cell in ws[row]):
            ws.delete_rows(row)

    # Записываем новые итоги без лишней строки
    ws.append(["Итого доходы:", "", "", total_income])
    ws.append(["Итого расходы:", "", "", total_expense])
    ws.append(["Остаток:", "", "", balance])

    # Сделаем итоги жирными
    for i in range(3):  # последние 3 строки
        for cell in ws[ws.max_row - i]:
            cell.font = Font(bold=True)

    # Сохраняем файл
    wb.save(filename)

    # Печатаем результат в консоль
    print(f"\n✅ Данные успешно добавлены в файл '{filename}'!")
    print(f"📊 Доходы: {total_income} | Расходы: {total_expense} | Остаток: {balance}")

# 🔸 Запуск программы
entries = user_input()
if entries:
    income_expenses(entries)
else:
    print("❗ Данные не были введены.")