# Выгрузка assets  с конкретными уязвимостями и перечнем установленного ПО из данных, ролученных
# для конкретного предприятия через API KUMA

import json
import pandas as pd

# Загрузка данных из JSON файла
with open('epv.json', 'r', encoding='utf-8-sig') as file:
    data = json.load(file)

# Преобразование JSON данных в DataFrame
df = pd.json_normalize(data, sep='_')

# Создание DataFrame для установленного ПО
software_df = pd.json_normalize(
    data,
    record_path=['software'],
    meta=['id', 'name', 'fqdn'],  # Добавляем основные данные для связи
    meta_prefix='main_',
    sep='_'
)

# Создание DataFrame для уязвимостей
vulnerabilities_df = pd.json_normalize(
    data,
    record_path=['vulnerabilities'],
    meta=['id', 'name', 'fqdn'],  # Добавляем основные данные для связи
    meta_prefix='main_',
    sep='_'
)

# Создание Excel файла с несколькими листами
with pd.ExcelWriter('epv.xlsx', engine='openpyxl') as writer:
    # Основные данные
    df.to_excel(writer, sheet_name='Получено через API', index=False)
    
    # Установленное ПО
    software_df.to_excel(writer, sheet_name='Установленное ПО', index=False)
    
    # Уязвимости
    vulnerabilities_df.to_excel(writer, sheet_name='Уязвимости', index=False)

print("Данные успешно загружены в Excel файл 'assets.xlsx'")