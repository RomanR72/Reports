import os
import re
import pandas as pd
from openpyxl import load_workbook
import tempfile

# Функция для извлечения кодов правил из названий
def extract_rule_code(rule_name):
    match = re.search(r'R\d+(_\d+)*', rule_name)
    return match.group(0) if match else None

# Функция для поиска тактики MITRE по коду правила
def find_mitre_tactic(rule_code, mitre_data):
    if not rule_code:
        return None
        
    code_only = rule_code[1:] if rule_code.startswith('R') else rule_code
    
    for col in mitre_data.columns[1:]:  # Пропускаем первый столбец
        if code_only in mitre_data[col].values:
            return mitre_data.columns[mitre_data.columns.get_loc(col)]
    return None

# Основной код
def main():
    try:
        # Загрузка данных MITRE
        mitre_file = 'MITRE.xlsx'
        mitre_data = pd.read_excel(mitre_file, sheet_name='Лист1', header=0)
        
        # Поиск processed-файлов в каталоге output
        output_dir = 'output'
        processed_files = [f for f in os.listdir(output_dir) 
                          if f.startswith('processed') and f.endswith('.xlsx')]
        
        # Сбор уникальных значений из столбца A листа 1-6
        unique_rules = set()
        
        for file in processed_files:
            file_path = os.path.join(output_dir, file)
            try:
                df = pd.read_excel(file_path, sheet_name='1-6', header=None)
                # Пропускаем первые две строки и берем данные из столбца A
                rules = df.iloc[2:, 0].dropna().unique()
                unique_rules.update(rules)
            except Exception as e:
                print(f"Ошибка при обработке файла {file}: {e}")
        
        # Создание DataFrame для результатов
        results = []
        
        for rule in unique_rules:
            rule_code = extract_rule_code(str(rule))  # Преобразуем в строку на всякий случай
            if rule_code:
                code_only = rule_code[1:]  # Убираем 'R'
                tactic = find_mitre_tactic(rule_code, mitre_data)
                results.append([rule, rule_code, code_only, tactic])
        
        # Создание итогового DataFrame
        result_df = pd.DataFrame(results, columns=[
            'Original_Rule', 'Rule_Code', 'Code_Only', 'MITRE_Tactic'
        ])
        
        # Попытка сохранения в текущую директорию
        try:
            result_df.to_excel('rules.xlsx', index=False)
            print("Файл rules.xlsx успешно создан в текущей директории")
        except Exception as e:
            print(f"Не удалось сохранить в текущую директорию: {e}")
            
            # Попытка сохранения в временную директорию
            try:
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, 'rules.xlsx')
                result_df.to_excel(temp_path, index=False)
                print(f"Файл сохранен во временную директорию: {temp_path}")
            except Exception as e2:
                print(f"Не удалось сохранить во временную директорию: {e2}")
                
                # Попытка сохранения в домашнюю директорию пользователя
                try:
                    home_dir = os.path.expanduser("~")
                    home_path = os.path.join(home_dir, 'rules.xlsx')
                    result_df.to_excel(home_path, index=False)
                    print(f"Файл сохранен в домашнюю директорию: {home_path}")
                except Exception as e3:
                    print(f"Не удалось сохранить файл ни в одну из директорий: {e3}")
                    # Вывод данных в консоль как последний вариант
                    print("Данные для сохранения:")
                    print(result_df.to_string())
                    
    except Exception as e:
        print(f"Общая ошибка выполнения: {e}")

if __name__ == "__main__":
    main()