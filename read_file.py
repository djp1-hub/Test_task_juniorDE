import pandas as pd
from datetime import datetime

file_path = 'random_sales_data_with_tables_and_styles.xlsx'
start_time = datetime.now()

# Чтение всех листов в один словарь DataFrames
excel_sheets = pd.read_excel(file_path, sheet_name=None)

# Дополнительное (и лишнее) чтение каждого листа
# Основное объединение всех листов
united_df = pd.concat([df for df in excel_sheets.values()])

# Лишний расчет времени для каждой части отдельно
end_time_intermediate = datetime.now()
print(f'Intermediate Duration: {(end_time_intermediate - start_time).total_seconds()} seconds')

# Группировка по дате продажи и суммирование количества
final_start_time = datetime.now()
grouped_df = united_df.groupby("Дата продажи")["Количество"].sum()
final_end_time = datetime.now()

# Общий расчет времени выполнения
print(grouped_df)
duration = datetime.now() - start_time
print(f'Total Duration: {duration.total_seconds()} seconds')
print(f'Number of rows in united DataFrame: {len(united_df)}')