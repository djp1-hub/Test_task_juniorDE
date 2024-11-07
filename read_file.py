import re
import time
import zipfile
from functools import wraps

import pandas as pd
from datetime import datetime, timedelta
import xml.etree.ElementTree as ET

data_types = {
    "Дата продажи": "datetime64[ns]",
    "ID товара": int,
    "Название товара": str,
    "Категория товара": str,
    "Бренд": str,
    "Цена за единицу": float,
    "Количество": int,
    "Сумма продажи": float,
    "Скидка (%)": float,
    "Сумма скидки": float,
    "Итоговая сумма": float,
    "ID клиента": int,
    "Имя клиента": str,
    "Регион": str,
    "Город": str,
    "Метод оплаты": str,
    "Канал продажи": str,
    "Тип доставки": str,
    "Статус доставки": str,
    "Дата доставки": "datetime64[ns]",
    "Менеджер по продажам": str,
    "Отдел продаж": str,
    "Месяц": int,
    "Год": int
}

def timer(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.perf_counter()
        result = func(*args, **kwargs)
        end_time = time.perf_counter()
        print(f"Функция {func.__name__} выполнялась {end_time - start_time:.4f} секунд")
        return result
    return wrapper


@timer
def old_version(file_path: str) -> pd.DataFrame:
    # Чтение всех листов в один словарь DataFrames
    excel_sheets = pd.read_excel(file_path, sheet_name=None)

    # Основное объединение всех листов
    united_df = pd.concat([df for df in excel_sheets.values()])

    # Группировка по дате продажи и суммирование количества
    grouped_df = united_df.groupby("Дата продажи")["Количество"].sum()

    return grouped_df


def normalize_date(excel_date):
    return datetime(1899, 12, 30) + timedelta(days=int(excel_date))


# Чтение данных из xml в массив
def xml_parser(text) -> []:
    root = ET.fromstring(text)
    namespace = {'ns': root.tag.split('}')[0].strip('{')}
    sheet_data = root.find('ns:sheetData', namespace)
    rows = []

    for row in sheet_data:
        row_data = []

        for i in range(len(row)):
            if row[i].attrib.get('t') == 'inlineStr':
                value = str(row[i].find('ns:is/ns:t', namespace).text)
            elif row[i].attrib.get('t') == 'n':
                dtype = data_types.get(rows[0][i])
                if dtype == "datetime64[ns]":
                    value = normalize_date(row[i].find('ns:v', namespace).text)
                else:
                    value = dtype(row[i].find('ns:v', namespace).text)
            else:
                value = None

            row_data.append(value)
        rows.append(row_data)

    return rows


@timer
def my_version(file_path: str, pattern:str) -> pd.DataFrame:
    # чтение данных из файла .xlsx
    with zipfile.ZipFile(file_path, 'r') as zf:
        path_list = zf.namelist()
        full_data = []
        sheet_data = []

        for line in path_list:
            if re.match(pattern, line):
                with zf.open(line, 'r') as file:
                    sheet_data = xml_parser(file.read())
                    full_data += sheet_data[1:]

    # Приведение данных к DataFrame
    df = pd.DataFrame(full_data, columns=sheet_data[0])

    # Группировка по дате продажи и суммирование количества
    grouped_df = df.groupby("Дата продажи")["Количество"].sum()

    return grouped_df


file_path = 'random_sales_data_with_tables_and_styles.xlsx'
pattern = r"xl/worksheets/sheet.*\.xml"

my_result = my_version(file_path, pattern)
old_result = old_version(file_path)

print(my_result.equals(old_result))