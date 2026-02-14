"""
Парсер файла продаж заказных и незаказных товаров
Структура: два блока "Незаказной" и "Заказной"
Колонки: D - количество, F - продажи, E - прибыль
"""

import pandas as pd
import re
import os

def normalize_fio(name):
    """Нормализует ФИО для сравнения"""
    if not name or not isinstance(name, str):
        return ''
    return ' '.join(name.strip().split()).upper()

def parse_zakaz_number(value):
    """Парсит число из строки, убирает пробелы, заменяет запятую на точку"""
    if not value or not isinstance(value, str):
        return 0
    cleaned = value.replace(' ', '').replace(',', '.').strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0


def parse_zakaz_sales(filepath, staff_data=None, excluded_firms=None):
    """
    Парсит файл Заказ.xls, ищет только сотрудников из staff_data
    
    Args:
        filepath: путь к файлу Заказ.xls
        staff_data: данные сотрудников {'ФИО_норм': 'ФИО'}
        excluded_firms: список фирм для исключения из УРС
    
    Returns:
        dict: данные по найденным сотрудникам
    """
    result = {
        'success': False,
        'data': {},
        'statistics': {
            'total_unordered_items': 0,
            'total_unordered_revenue': 0,
            'total_unordered_profit': 0,
            'total_ordered_items': 0,
            'total_ordered_revenue': 0,
            'total_ordered_profit': 0,
            'vendors_count': 0,
            'matched_employees': 0
        },
        'error': None
    }
    
    def is_valid_seller_name(name, excluded_firms_list):
        """
        Проверяет, является ли строка валидным ФИО продавца
        """
        if excluded_firms_list is None:
            excluded_firms_list = []
        
        if not name or not isinstance(name, str):
            return False
        
        name_clean = name.strip()
        name_lower = name_clean.lower()
        
        if len(name_clean) < 4:
            return False
        
        # 1. ТОЧНОЕ СОВПАДЕНИЕ С ИСКЛЮЧЕНИЯМИ
        for exclusion in excluded_firms_list:
            if not exclusion or not isinstance(exclusion, str):
                continue
            exclusion_lower = exclusion.lower().strip()
            
            # Точное совпадение
            if name_lower == exclusion_lower:
                return False
        
        # 2. СОДЕРЖИТ КЛЮЧЕВЫЕ СЛОВА ИСКЛЮЧЕНИЙ
        for exclusion in excluded_firms_list:
            if not exclusion or not isinstance(exclusion, str):
                continue
            exclusion_lower = exclusion.lower().strip()
            
            # Частичное совпадение (если исключение содержится в имени)
            if exclusion_lower and exclusion_lower in name_lower:
                return False
        
        # 3. ПРОВЕРКА НА ТИПОВЫЕ ЗАГОЛОВКИ И ОБОБЩЕНИЯ
        invalid_keywords = [
            'незаказной', 'заказной', 'товар', 'продавец',
            'итого', 'всего', 'итог', 'общий', 'основной',
            '%', 'процент', 'руб.', 'рублей', 'ед.'
        ]
        
        for keyword in invalid_keywords:
            if keyword in name_lower:
                return False
        
        # 4. ПРОВЕРКА НА ЧИСЛА (не ФИО)
        if name_clean.replace(' ', '').isdigit():
            return False
        
        # 5. ПРОВЕРКА НА ПУСТЫЕ ИЛИ СЛИШКОМ КОРОТКИЕ
        if len(name_clean) < 2:
            return False
        
        # 6. ПРОВЕРКА НА РУССКИЕ БУКВЫ (опционально)
        has_cyrillic = any('а' <= char <= 'я' or char == 'ё' for char in name_lower)
        if not has_cyrillic:
            return False

        # 7. ПРОВЕРКА НА ФОРМАТ "ФАМИЛИЯ И.О." (инициалы через точку)
        if '.' in name_lower and len(name_clean.split()) <= 2:
            return False
        
        return True
    
    try:
        if not os.path.exists(filepath):
            result['error'] = f"Файл не найден: {filepath}"
            return result
        
        print(f"📊 Парсим файл: {os.path.basename(filepath)}")
        
        # Получаем список ФИО сотрудников для поиска
        employee_names = {}
        if staff_data and staff_data.get('success'):
            for emp in staff_data.get('employees', []):
                fio_norm = emp.get('ФИО_норм', '')
                fio_original = emp.get('ФИО', '')
                if fio_norm:
                    employee_names[fio_norm] = fio_original
            print(f"🔍 Ищем {len(employee_names)} сотрудников из staff_data")
            print(f"DEBUG: Первые 5 сотрудников: {list(employee_names.keys())[:5]}")
        
        # Читаем файл
        if filepath.endswith('.xls'):
            df = pd.read_excel(filepath, header=None, engine='xlrd')
        else:
            df = pd.read_excel(filepath, header=None, engine='openpyxl')
        
        print(f"📄 Размер файла: {df.shape[0]} строк, {df.shape[1]} колонок")
        
        # Подготовка
        current_section = None
        vendors_data = {}
        matched_count = 0
        
        # Поиск данных
        for idx in range(len(df)):
            row = df.iloc[idx]
            
            # Проверяем начало секций
            cell0 = str(row[0]).strip() if pd.notnull(row[0]) else ''
            
            if 'Незаказной' in cell0 and 'товар' not in cell0.lower():
                current_section = 'unordered'
                continue
            elif 'Заказной' in cell0 and 'товар' not in cell0.lower():
                current_section = 'ordered'
                continue
            
            # Пропускаем если не в секции
            if not current_section:
                continue
            
            # Пропускаем пустые строки
            if row.isnull().all() or not cell0:
                continue
            
            # Проверяем валидность имени продавца
            if not is_valid_seller_name(cell0, excluded_firms):
                continue

            # Нормализуем имя из файла
            vendor_norm = normalize_fio(cell0)
            print(f"DEBUG: Имя из файла: '{cell0}' -> нормализовано: '{vendor_norm}'")
            
            # Ищем соответствие с сотрудниками
            matched_employee = None
            for emp_norm, emp_original in employee_names.items():
                # Простое сравнение нормализованных имен
                if emp_norm == vendor_norm:
                    matched_employee = emp_norm
                    break
                # Частичное совпадение (если полное не сработало)
                elif emp_norm in vendor_norm or vendor_norm in emp_norm:
                    matched_employee = emp_norm
                    print(f"  🔍 Частичное совпадение: '{emp_norm}' → '{vendor_norm}'")
                    break
            
            # Если не нашли сотрудника - пропускаем
            if not matched_employee:
                continue
            
            # Получаем данные
            if len(row) >= 7:
                # Колонки: D(3)=кол-во, F(5)=продажи, G(6)=прибыль
                items = parse_zakaz_number(str(row[3])) if pd.notnull(row[3]) else 0
                revenue = parse_zakaz_number(str(row[5])) if pd.notnull(row[5]) else 0
                profit = parse_zakaz_number(str(row[6])) if pd.notnull(row[6]) else 0
                
                # Создаем запись
                if matched_employee not in vendors_data:
                    vendors_data[matched_employee] = {
                        'fio': employee_names[matched_employee],
                        'unordered': {'items': 0, 'revenue': 0, 'profit': 0},
                        'ordered': {'items': 0, 'revenue': 0, 'profit': 0}
                    }
                    matched_count += 1
                
                # Добавляем данные
                vendors_data[matched_employee][current_section] = {
                    'items': items,
                    'revenue': revenue,
                    'profit': profit
                }
                
                # Статистика
                if current_section == 'unordered':
                    result['statistics']['total_unordered_items'] += items
                    result['statistics']['total_unordered_revenue'] += revenue
                    result['statistics']['total_unordered_profit'] += profit
                else:
                    result['statistics']['total_ordered_items'] += items
                    result['statistics']['total_ordered_revenue'] += revenue
                    result['statistics']['total_ordered_profit'] += profit
                print(f"DEBUG: Первые 5 сотрудников из staff_data: {list(employee_names.keys())[:5]}")
        
        result['data'] = vendors_data
        result['statistics']['vendors_count'] = len(vendors_data)
        result['statistics']['matched_employees'] = matched_count
        result['success'] = True
        
        print(f"✅ Парсинг завершен:")
        print(f"   • Найдено сотрудников: {matched_count} из {len(employee_names)}")
        print(f"   • Незаказные товары: {result['statistics']['total_unordered_items']:.0f} ед.")
        print(f"   • Заказные товары: {result['statistics']['total_ordered_items']:.0f} ед.")
        
        if matched_count > 0:
            print(f"  🔍 Примеры найденных сотрудников:")
            for i, (emp_norm, data) in enumerate(list(vendors_data.items())[:3], 1):
                print(f"     {i}. {data['fio']}:")
                print(f"        Незаказные: {data['unordered']['items']:.0f} ед., {data['unordered']['profit']:.0f} руб.")
                print(f"        Заказные: {data['ordered']['items']:.0f} ед., {data['ordered']['profit']:.0f} руб.")
        
    except Exception as e:
        result['error'] = f"Ошибка парсинга: {str(e)}"
        import traceback
        print(f"❌ Ошибка: {traceback.format_exc()}")
    
    return result
