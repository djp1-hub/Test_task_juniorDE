import pandas as pd
from datetime import datetime

file_path = 'random_sales_data_with_tables_and_styles.xlsx'
start_time = datetime.now()

# Чтение всех листов в один словарь DataFrames
excel_sheets = pd.read_excel(file_path, sheet_name=None)

# Основное объединение всех листов
united_df = pd.concat([df for df in excel_sheets.values()])


# Группировка по дате продажи и суммирование количества
grouped_df = united_df.groupby("Дата продажи")["Количество"].sum()

# Общий расчет времени выполнения
print(grouped_df)
duration = datetime.now() - start_time
print(f'Total Duration: {duration.total_seconds()} seconds')
print(f'Number of rows in united DataFrame: {len(united_df)}')
