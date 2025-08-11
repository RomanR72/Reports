import json
import pandas as pd

# Загрузка данных из JSON файла
with open('epv.json', 'r', encoding='utf-8-sig') as file:
    data = json.load(file)

# Преобразование JSON данных в DataFrame
df = pd.json_normalize(data, sep='_')

# Сохранение DataFrame в Excel файл
df.to_excel('output.xlsx', index=True, engine='openpyxl')

print("Данные успешно загружены в Excel файл 'output.xlsx'")
print("everything works")