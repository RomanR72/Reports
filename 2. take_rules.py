import os
import openpyxl
from openpyxl import Workbook
import re


def has_three_digits(s):
    """Проверяет, содержит ли строка ровно 3 цифры"""
    if not s:
        return False
    digits = re.findall(r'\d', str(s))
    return len(digits) == 3

def starts_with_p_and_digits(s):
    """Проверяет, начинается ли строка с P и цифр"""
    if not s:
        return False
    return bool(re.match(r'^P\d', str(s), re.IGNORECASE))

def process_files():
    # Получаем путь к директории со скриптом
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Создаем новый файл rules.xlsx
    output_wb = Workbook()
    output_ws = output_wb.active
    output_ws.title = "Processed Data"
    
    # Устанавливаем заголовки столбцов
    headers = ["Original", "First 4 chars", "Digits", "MITRE Match"]
    for col, header in enumerate(headers, start=1):
        output_ws.cell(row=1, column=col, value=header)
    
    # Множество для хранения уникальных значений
    unique_values = set()
    
    # Проверяем существование каталога output
    output_dir = os.path.join(script_dir, 'output')
    if not os.path.exists(output_dir):
        print(f"Ошибка: Каталог {output_dir} не существует!")
        return
    
    # Загружаем данные из MITRE.xlsx
    mitre_data = {}
    mitre_path = os.path.join(script_dir, 'MITRE.xlsx')
    if os.path.exists(mitre_path):
        try:
            mitre_wb = openpyxl.load_workbook(mitre_path)
            mitre_ws = mitre_wb.active
            
            # Собираем данные из всех столбцов MITRE.xlsx
            for col in mitre_ws.iter_cols():
                header = col[0].value if col[0].value else ""
                for cell in col[1:]:
                    if cell.value:
                        # Приводим значение к строке и удаляем лишние пробелы
                        cell_value = str(cell.value).strip()
                        mitre_data[cell_value] = header
        except Exception as e:
            print(f"Ошибка при чтении MITRE.xlsx: {e}")
    else:
        print(f"Предупреждение: Файл MITRE.xlsx не найден в {script_dir}")
    
    # Обрабатываем файлы в каталоге output
    processed_files = 0
    for filename in os.listdir(output_dir):
        if filename.endswith('.xlsx') and filename != 'rules.xlsx':
            filepath = os.path.join(output_dir, filename)
            try:
                wb = openpyxl.load_workbook(filepath)
                if '1-6' in wb.sheetnames:
                    sheet = wb['1-6']
                    
                    # Читаем значения из столбца A, начиная с A3
                    for row in sheet.iter_rows(min_row=3, min_col=1, max_col=1):
                        cell_value = row[0].value
                        
                        if cell_value and isinstance(cell_value, str):
                            cell_value = cell_value.strip()
                            
                            # Проверяем условие: если есть _, то после него не должно быть цифр
                            if '_' in cell_value:
                                parts = cell_value.split('_')
                                if len(parts) > 1 and any(c.isdigit() for c in parts[1]):
                                    continue
                            
                            # Добавляем только уникальные значения
                            if cell_value not in unique_values:
                                unique_values.add(cell_value)
                                processed_files += 1
            except Exception as e:
                print(f"Ошибка при обработке файла {filename}: {e}")
    
    print(f"Обработано файлов: {processed_files}")
    print(f"Найдено уникальных значений: {len(unique_values)}")
    
    # Записываем данные в выходной файл с двойной фильтрацией
    rows_to_keep = [headers]  # Сохраняем заголовки
    for value in sorted(unique_values):
        # Проверяем условия:
        # 1. Содержит ровно 3 цифры
        # 2. Не начинается с P и цифр
        if has_three_digits(value) and not starts_with_p_and_digits(value):
            first_4 = value[:4] if value else ""
            digits = ''.join(filter(str.isdigit, first_4)) if first_4 else ""
            
            # Проверяем совпадение с MITRE.xlsx
            mitre_match = mitre_data.get(digits, "")
            
            rows_to_keep.append([value, first_4, digits if digits else None, mitre_match])
    
    # Очищаем лист и записываем только подходящие строки
    output_ws.delete_rows(1, output_ws.max_row)
    for row in rows_to_keep:
        output_ws.append(row)
    
    # Сохраняем файл rules.xlsx
    output_path = os.path.join(script_dir, 'rules.xlsx')
    try:
        output_wb.save(output_path)
        print(f"Файл успешно создан: {output_path}")
        print(f"Сохранено строк: {len(rows_to_keep)-1} (с 3 цифрами и не начинающихся с P+цифры)")
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")

if __name__ == "__main__":
    process_files()