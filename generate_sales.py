import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import string
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.protection import SheetProtection

# Генерация случайного пароля для защиты
password = ''.join(random.choices(string.ascii_letters, k=8))

# Определяем типы данных для столбцов
data_types = {
    "Дата продажи": "datetime",
    "ID товара": "int",
    "Название товара": "str",
    "Категория товара": "str",
    "Бренд": "str",
    "Цена за единицу": "float",
    "Количество": "int",
    "Сумма продажи": "float",
    "Скидка (%)": "float",
    "Сумма скидки": "float",
    "Итоговая сумма": "float",
    "ID клиента": "int",
    "Имя клиента": "str",
    "Регион": "str",
    "Город": "str",
    "Метод оплаты": "str",
    "Канал продажи": "str",
    "Тип доставки": "str",
    "Статус доставки": "str",
    "Дата доставки": "datetime",
    "Менеджер по продажам": "str",
    "Отдел продаж": "str",
    "Месяц": "int",
    "Год": "int"
}


def random_date(start_date):
    """Генерирует случайную дату в пределах года от заданной даты."""
    return start_date + timedelta(days=random.randint(0, 365))


def generate_data(row_count=30):
    """Генерирует DataFrame со случайными данными на основе структуры data_types."""
    data = {}
    for col, dtype in data_types.items():
        if dtype == "datetime":
            data[col] = [random_date(datetime(2023, 1, 1)) for _ in range(row_count)]
        elif dtype == "int":
            data[col] = np.random.randint(1, 1000, row_count).tolist()
        elif dtype == "float":
            data[col] = np.round(np.random.uniform(1.0, 1000.0, row_count), 2).tolist()
        elif dtype == "str":
            data[col] = [f"{col}_{random.randint(1, 100)}" for _ in range(row_count)]
    return pd.DataFrame(data)


# Генерация Excel файла с 335 листами
excel_path = "random_sales_data_with_tables_and_styles.xlsx"
with pd.ExcelWriter(excel_path) as writer:
    for i in range(335):
        sheet_data = generate_data()
        sheet_data.to_excel(writer, sheet_name=f"Sheet_{i}", index=False)

# Загрузка сгенерированного Excel файла
workbook = load_workbook(excel_path)


def random_style_name():
    """Создает случайное имя стиля."""
    return ''.join(random.choices(string.ascii_letters, k=42))


# Создание и добавление уникальных стилей
style_names = []
for i in range(10):
    style_name = random_style_name()
    style = NamedStyle(name=style_name)
    style.font = Font(size=10 + (i % 5), bold=(i % 2 == 0))
    style.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                          top=Side(style="thin"), bottom=Side(style="thin"))
    workbook.add_named_style(style)
    style_names.append(style_name)


def random_color():
    """Генерирует случайный цвет в формате RGB."""
    return f"{random.randint(225, 255):02X}{random.randint(225, 255):02X}{random.randint(225, 255):02X}"


# Генерация списка уникальных цветов
unique_colors = [random_color() for _ in range(35)]

# Применение таблиц, стилей и цветов к каждой вкладке
for sheet_name in workbook.sheetnames:
    worksheet = workbook[sheet_name]

    end_column = worksheet.max_column
    end_row = worksheet.max_row
    table_ref = f"A1:{chr(64 + end_column)}{end_row}"

    # Создание и добавление таблицы с заданным стилем
    table = Table(displayName=sheet_name, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9", showFirstColumn=False,
        showLastColumn=False, showRowStripes=True, showColumnStripes=True
    )
    worksheet.add_table(table)

    # Применение стилей и цветов к каждой ячейке
    for row in range(1, end_row + 1):
        for col in range(1, end_column + 1):
            cell = worksheet.cell(row=row, column=col)
            cell.style = random.choice(style_names)
            cell.fill = PatternFill(start_color=random.choice(unique_colors), end_color=random.choice(unique_colors),
                                    fill_type="solid")

    # Форматирование столбца "Дата продажи" с датой
    for row in range(2, worksheet.max_row + 1):  # начиная со второй строки
        cell = worksheet.cell(row=row, column=1)
        if isinstance(cell.value, datetime):
            cell.number_format = 'DD.MM.YYYY'

    # Установка защиты листа с паролем
    worksheet.protection = SheetProtection(sheet=True, password=password)

# Сохранение книги
workbook.save(excel_path)
