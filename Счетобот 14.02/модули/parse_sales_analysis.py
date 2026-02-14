# parse_sales_analysis.py
import pandas as pd
import re

def is_valid_seller_name(name, excluded_firms=None):
    """
    Проверяет, является ли строка валидным ФИО продавца
    
    name: строка для проверки
    excluded_firms: список названий фирм для исключения
    """
    if excluded_firms is None:
        excluded_firms = []
    
    if not name or not isinstance(name, str):
        return False
    
    name_clean = name.strip()
    name_lower = name_clean.lower()
    
    if len(name_clean) < 4:
        return False
    
    # 1. ТОЧНОЕ СОВПАДЕНИЕ С ИСКЛЮЧЕНИЯМИ
    for exclusion in excluded_firms:
        if not exclusion or not isinstance(exclusion, str):
            continue
        exclusion_lower = exclusion.lower().strip()
        
        # Точное совпадение
        if name_lower == exclusion_lower:
            return False
    
    # 2. ЗАПРЕЩЁННЫЕ СЛОВА
    forbidden_patterns = [
        'итого', 'всего', 'БД1', 'наименование', 'БД3', 'компания',
        'оптовая', 'розничная', 'продажа', 'БД4', 'прочая', 'отдел',
        'подразделение', 'филиал', 'управление', 'департамент',
        '!!!!', 'nan', 'none',
        'оптова', 'розничн', 'по чек', 'прочая', 'керамика',
        'сантехника', 'инструмент', 'отпуск', 'в отпуске', 'болен',
        'больничный', 'самообслуж', 'монтаж', 'ламинат', 'обои ',
        ' обои', 'паркет', 'электр', 'продаж'
    ]
    
    # Проверяем, не является ли строка типом продаж (содержит "продажа" но не только это)
    if 'продажа' in name_lower and len(name_clean.split()) <= 3:
        return False
    
    for pattern in forbidden_patterns:
        if pattern in name_lower:
            return False
    
    # 3. ПРОВЕРКА НА ФИО
    words = name_clean.split()
    if len(words) < 2:
        return False
    
    russian_letters = set('абвгдеёжзийклмнопрстуфхцчшщъыьэюя')
    for word in words:
        has_russian = any(c.lower() in russian_letters for c in word)
        if not has_russian:
            return False
    
    if name_clean.isupper() and len(name_clean) > 20:
        return False
    
    if any(char.isdigit() for char in name_clean):
        return False
    
    if any(symbol in name_clean for symbol in ['"', '«', '»', '()', 'ООО', 'ИП', 'АО', 'ЗАО']):
        return False
    
    return True

def parse_sales_analysis(file_path, bonus_items_set, non_liquid_items_set, exclusions=None):
    """
    Парсер файла 'Анализ продаж' для структуры:
    A: ФИО/фирма/тип/код | B: Наименование | C: Ед. | D: Кол | E: Себестоимость | F: Продажи | G: Прибыль
    
    Возвращает словарь с детализацией по продавцам, включая:
    - Все продажи (выручка, прибыль, количество)
    - Продажи по типам (опт, розница по чекам, розница прочая)
    - Бонусные товары (количество, выручка, прибыль)
    - Неликвидные товары (количество, выручка, прибыль=0)
    """
    print(f"📊 Парсинг файла продаж: {file_path}")
    
    if exclusions is None:
        exclusions = []
    
    print(f"  🚫 Исключений из УРС: {len(exclusions)}")
    
    try:
        # Читаем файл
        df = pd.read_excel(file_path, header=None, dtype=str)
        
        # Получаем период из ячейки B5
        period_cell = ""
        if len(df) > 4 and df.shape[1] > 1:
            period_cell = str(df.iloc[4, 1]).strip()
            print(f"📅 Период: {period_cell}")
        
        # Ищем начало таблицы с продажами
        start_row = None
        
        # Ищем строку с "Продавец" или "ФИО" в колонке A
        for i in range(5, min(50, len(df))):
            cell_a = str(df.iloc[i, 0]).strip().lower() if df.shape[1] > 0 else ""
            
            if any(word in cell_a for word in ['продавец', 'фио', 'сотрудник', 'менеджер']):
                start_row = i
                print(f"🔍 Найден заголовок таблицы: строка {start_row + 1}")
                print(f"   Содержимое: '{str(df.iloc[i, 0]).strip()}'")
                break
        
        if start_row is None:
            # Если не нашли по заголовку, ищем первую строку с валидным ФИО
            for i in range(5, min(100, len(df))):
                cell_a = str(df.iloc[i, 0]).strip()
                if is_valid_seller_name(cell_a, exclusions):
                    start_row = i - 1
                    print(f"🔍 Найден первый продавец: строка {i + 1}")
                    print(f"   Устанавливаем начало таблицы: строка {start_row + 1}")
                    break
        
        if start_row is None:
            start_row = 5
            print(f"⚠️  Не найден заголовок, начинаем со строки {start_row + 1}")
        
        # ОПРЕДЕЛЯЕМ КОЛОНКИ ПО ФИКСИРОВАННОЙ СТРУКТУРЕ
        col_mapping = {
            'фио': 0,        # Колонка A: ФИО/фирма/тип/код
            'наименование': 1, # Колонка B: Наименование товара
            'единица': 2,    # Колонка C: Ед.
            'количество': 3,  # Колонка D: Кол
            'себестоимость': 4, # Колонка E: Себестоимость
            'выручка': 5,    # Колонка F: Продажи
            'прибыль': 6     # Колонка G: Прибыль
        }
        
        print(f"📋 Структура колонок (фиксированная):")
        print(f"  A (0): ФИО/фирма/тип/код")
        print(f"  B (1): Наименование товара")
        print(f"  C (2): Единица измерения")
        print(f"  D (3): Количество")
        print(f"  E (4): Себестоимость")
        print(f"  F (5): Продажи (выручка)")
        print(f"  G (6): Прибыль")
        
        # Проверяем, что файл имеет достаточно колонок
        if df.shape[1] < 7:
            print(f"❌ ОШИБКА: Файл имеет только {df.shape[1]} колонок, нужно минимум 7")
            return {}
        
        # Парсим данные с учетом иерархии
        sales_data = {}
        
        # Переменные для отслеживания текущего состояния
        current_seller = None
        current_seller_normalized = None
        current_sale_type = None
        in_seller_block = False
        
        valid_sellers = 0
        items_count_total = 0
        
        print(f"\n🔍 Начинаю парсинг данных со строки {start_row + 1}...")
        
        for i in range(start_row + 1, len(df)):
            # Получаем данные из колонки A
            cell_a = str(df.iloc[i, 0]).strip() if df.shape[1] > 0 else ""
            
            # Пропускаем пустые строки
            if not cell_a or cell_a.lower() in ['', 'nan', 'none']:
                if in_seller_block:
                    in_seller_block = False
                    current_seller = None
                    current_seller_normalized = None
                    current_sale_type = None
                continue
            
            # Проверяем, является ли это валидным ФИО продавца
            if is_valid_seller_name(cell_a, exclusions):
                # НАЧАЛО НОВОГО ПРОДАВЦА
                fio = cell_a
                fio_normalized = ' '.join(fio.split()).upper()
                current_seller = fio
                current_seller_normalized = fio_normalized
                current_sale_type = None
                in_seller_block = True
                valid_sellers += 1
                
                # Инициализируем нового продавца с ДЕТАЛИЗАЦИЕЙ
                if fio_normalized not in sales_data:
                    sales_data[fio_normalized] = {
                        'department': "Не указан",
                        'sales_by_type': {
                            'Оптовая продажа': {
                                'revenue': 0.0, 'profit': 0.0, 'items_count': 0,
                                'bonus_items_count': 0, 'bonus_revenue': 0.0, 'bonus_profit': 0.0,
                                'non_liquid_items_count': 0, 'non_liquid_revenue': 0.0, 'non_liquid_profit': 0.0,
                                'regular_profit': 0.0, 'regular_revenue': 0.0  # ← ДОБАВЬ
                            },
                            'Розничная (по чекам)': {
                                'revenue': 0.0, 'profit': 0.0, 'items_count': 0,
                                'bonus_items_count': 0, 'bonus_revenue': 0.0, 'bonus_profit': 0.0,
                                'non_liquid_items_count': 0, 'non_liquid_revenue': 0.0, 'non_liquid_profit': 0.0,
                                'regular_profit': 0.0, 'regular_revenue': 0.0  # ← ДОБАВЬ
                            },
                            'Розничная (прочая)': {
                                'revenue': 0.0, 'profit': 0.0, 'items_count': 0,
                                'bonus_items_count': 0, 'bonus_revenue': 0.0, 'bonus_profit': 0.0,
                                'non_liquid_items_count': 0, 'non_liquid_revenue': 0.0, 'non_liquid_profit': 0.0,
                                'regular_profit': 0.0, 'regular_revenue': 0.0  # ← ДОБАВЬ
                            }
                        },
                        'total_revenue': 0.0,
                        'total_profit': 0.0,
                        'total_items_count': 0,
                        'total_bonus_items_count': 0,
                        'total_bonus_revenue': 0.0,
                        'total_bonus_profit': 0.0,
                        'total_non_liquid_items_count': 0,
                        'total_non_liquid_revenue': 0.0,
                        'total_non_liquid_profit': 0.0,
                        'original_name': fio,
                        'row_number': i + 1
                    }
                
                if valid_sellers <= 3:
                    print(f"  ✅ Найден продавец {valid_sellers}: '{fio}' (строка {i + 1})")
                
                continue
            
            # Если мы в блоке продавца
            if in_seller_block and current_seller_normalized and current_seller_normalized in sales_data:
                
                # Проверяем, является ли строка типом продаж
                cell_a_lower = cell_a.lower()
                if any(sale_type in cell_a_lower for sale_type in ['оптовая', 'розничная', 'продажа']):
                    # Определяем тип продаж
                    if 'оптовая' in cell_a_lower:
                        current_sale_type = 'Оптовая продажа'
                        if valid_sellers <= 3:
                            print(f"    → Тип продаж: 'Оптовая продажа'")
                    elif 'розничная' in cell_a_lower and 'по чек' in cell_a_lower:
                        current_sale_type = 'Розничная (по чекам)'
                    elif 'розничная' in cell_a_lower:
                        current_sale_type = 'Розничная (прочая)'
                    else:
                        current_sale_type = None
                    continue
                
                # Проверяем, является ли строка товаром (начинается с цифрового кода)
                clean_cell_a = cell_a.replace(' ', '').replace('-', '').replace('.', '')
                if clean_cell_a.isdigit() and 3 <= len(clean_cell_a) <= 8:
                    # ЭТО ТОВАР - обрабатываем
                    item_code = clean_cell_a
                    
                    # Получаем количество единиц (колонка D)
                    items_count = 0
                    if df.shape[1] > 3:
                        qty_cell = str(df.iloc[i, 3]).strip()
                        if qty_cell and qty_cell.lower() not in ['', 'nan', 'none']:
                            try:
                                qty_cell_clean = qty_cell.replace(',', '.').replace(' ', '')
                                items_count = float(qty_cell_clean)
                            except:
                                items_count = 0
                    
                    # Получаем выручку (колонка F)
                    revenue = 0.0
                    if df.shape[1] > 5:
                        revenue_cell = str(df.iloc[i, 5]).strip()
                        if revenue_cell and revenue_cell.lower() not in ['', 'nan', 'none']:
                            try:
                                revenue_cell_clean = revenue_cell.replace(',', '.').replace(' ', '')
                                revenue = float(revenue_cell_clean)
                            except:
                                revenue = 0.0
                    
                    # Получаем прибыль (колонка G)
                    profit = 0.0
                    if df.shape[1] > 6:
                        profit_cell = str(df.iloc[i, 6]).strip()
                        if profit_cell and profit_cell.lower() not in ['', 'nan', 'none']:
                            try:
                                profit_cell_clean = profit_cell.replace(',', '.').replace(' ', '')
                                profit = float(profit_cell_clean)
                            except:
                                profit = 0.0
                    
                    # Определяем тип товара по коду
                    if item_code in bonus_items_set:
                        item_type = 'bonus'
                    elif item_code in non_liquid_items_set:
                        item_type = 'non_liquid'
                    else:
                        item_type = 'regular'
                    
                    # Если нет типа продаж, используем "Оптовая продажа" по умолчанию
                    if not current_sale_type:
                        current_sale_type = 'Оптовая продажа'
                    
                    # Получаем данные текущего типа продаж (ПОСЛЕ определения типа продаж!)
                    type_data = sales_data[current_seller_normalized]['sales_by_type'][current_sale_type]
                    
                    # Общие показатели типа продаж
                    type_data['items_count'] += items_count
                    type_data['revenue'] += revenue
                    type_data['profit'] += profit
                    
                    # РАЗДЕЛЕНИЕ ПО ТИПАМ ТОВАРОВ (ДОБАВЛЕНО)
                    if item_type == 'regular':
                        type_data['regular_revenue'] = type_data.get('regular_revenue', 0) + revenue
                        type_data['regular_profit'] = type_data.get('regular_profit', 0) + profit
                    
                    elif item_type == 'bonus':
                        type_data['bonus_items_count'] += items_count
                        type_data['bonus_revenue'] += revenue
                        type_data['bonus_profit'] += profit
                        
                        sales_data[current_seller_normalized]['total_bonus_items_count'] += items_count
                        sales_data[current_seller_normalized]['total_bonus_revenue'] += revenue
                        sales_data[current_seller_normalized]['total_bonus_profit'] += profit
                    
                    elif item_type == 'non_liquid':
                        type_data['non_liquid_items_count'] += items_count
                        type_data['non_liquid_revenue'] += revenue
                        type_data['non_liquid_profit'] += profit  # ← Теперь сохраняем прибыль неликвидов
                        
                        sales_data[current_seller_normalized]['total_non_liquid_items_count'] += items_count
                        sales_data[current_seller_normalized]['total_non_liquid_revenue'] += revenue
                        sales_data[current_seller_normalized]['total_non_liquid_profit'] += profit
                    
                    # Общие показатели продавца
                    sales_data[current_seller_normalized]['total_revenue'] += revenue
                    sales_data[current_seller_normalized]['total_profit'] += profit
                    sales_data[current_seller_normalized]['total_items_count'] += items_count
                    
                    items_count_total += items_count
                    
                    # Отладочный вывод для первых товаров
                    if items_count_total <= 10:
                        item_name = str(df.iloc[i, 1]).strip()[:30] if df.shape[1] > 1 else "нет названия"
                        type_text = {
                            'regular': 'обычный',
                            'bonus': 'БОНУС',
                            'non_liquid': 'НЕЛИКВИД'
                        }.get(item_type, '?')
                        print(f"    → Товар: {item_code} ({item_name}) - {items_count} шт. = {revenue:,.0f} руб. [{type_text}]")
                    
                    continue
            
            # Если строка содержит название фирмы или отдела (заглавные буквы)
            if cell_a.isupper() and len(cell_a) > 5:
                # Вероятно, это фирма или отдел - конец блока продавца
                in_seller_block = False
                current_seller = None
                current_seller_normalized = None
                current_sale_type = None
        
        # Статистика
        print(f"\n✅ СТАТИСТИКА ПАРСИНГА:")
        print(f"   • Валидных продавцов: {valid_sellers}")
        print(f"   • Уникальных продавцов: {len(sales_data)}")
        print(f"   • Общая выручка: {sum(s['total_revenue'] for s in sales_data.values()):,.0f} руб.")
        print(f"   • Всего единиц товара: {items_count_total:,.0f} шт.")
        print(f"   • Бонусных товаров: {sum(s['total_bonus_items_count'] for s in sales_data.values()):,.0f} шт.")
        print(f"   • Неликвидных товаров: {sum(s['total_non_liquid_items_count'] for s in sales_data.values()):,.0f} шт.")
        
        # Показываем примеры
        if sales_data:
            print(f"\n📋 ПРИМЕРЫ ПРОДАВЦОВ (первые 3):")
            for i, (key, data) in enumerate(list(sales_data.items())[:3], 1):
                name = data.get('original_name', key)
                revenue = data.get('total_revenue', 0)
                profit = data.get('total_profit', 0)
                bonus_items = data.get('total_bonus_items_count', 0)
                bonus_revenue = data.get('total_bonus_revenue', 0)
                non_liquid_items = data.get('total_non_liquid_items_count', 0)
                non_liquid_revenue = data.get('total_non_liquid_revenue', 0)
                
                print(f"\n  {i}. {name}")
                print(f"     Выручка: {revenue:,.0f} руб. | Прибыль: {profit:,.0f} руб.")
                print(f"     Бонусы: {bonus_items} шт. = {bonus_revenue:,.0f} руб.")
                print(f"     Неликвиды: {non_liquid_items} шт. = {non_liquid_revenue:,.0f} руб.")
                
                # Детали по типам продаж
                for sale_type, type_data in data['sales_by_type'].items():
                    if type_data['revenue'] > 0:
                        items = type_data['items_count']
                        type_revenue = type_data['revenue']
                        type_profit = type_data['profit']
                        bonus_count = type_data['bonus_items_count']
                        bonus_rev = type_data['bonus_revenue']
                        non_liquid_count = type_data['non_liquid_items_count']
                        non_liquid_rev = type_data['non_liquid_revenue']
                        
                        print(f"     • {sale_type}: {items} шт. = {type_revenue:,.0f} руб.")
                        if bonus_count > 0:
                            print(f"       Бонусы: {bonus_count} шт. = {bonus_rev:,.0f} руб.")
                        if non_liquid_count > 0:
                            print(f"       Неликвиды: {non_liquid_count} шт. = {non_liquid_rev:,.0f} руб.")
        else:
            print(f"\n⚠️  ВНИМАНИЕ: Нет данных о продавцах")
        
        # Преобразуем данные в формат для нового расчета
        print(f"\n🔄 Преобразование данных для нового расчета...")
        
        for seller_key, seller_data in sales_data.items():
            # Переименовываем поля для ясности
            sales_by_type = seller_data['sales_by_type']
            
            # 1. Розничная (по чекам) - для личного показателя
            розничная_чеки = sales_by_type.get('Розничная (по чекам)', {})
            seller_data['продажи_чеки'] = {
                'выручка': розничная_чеки.get('revenue', 0),
                'прибыль': розничная_чеки.get('profit', 0),
                'прибыль_обычная': розничная_чеки.get('regular_profit', 0),
                'прибыль_бонусная': розничная_чеки.get('bonus_profit', 0),
                'выручка_неликвидов': розничная_чеки.get('non_liquid_revenue', 0),
                'прибыль_неликвидов': розничная_чеки.get('non_liquid_profit', 0)
            }
            
            # 2. Оптовая продажа
            оптовая = sales_by_type.get('Оптовая продажа', {})
            seller_data['продажи_опт'] = {
                'прибыль': оптовая.get('profit', 0),
                'выручка': оптовая.get('revenue', 0),
                'items_count': оптовая.get('items_count', 0)
            }
            
            # 3. Розничная (прочая)
            розничная_прочая = sales_by_type.get('Розничная (прочая)', {})
            seller_data['продажи_прочая'] = {
                'выручка': розничная_прочая.get('revenue', 0),
                'прибыль': розничная_прочая.get('profit', 0),
                'items_count': розничная_прочая.get('items_count', 0)
            }
            
            # Итоговые поля для обратной совместимости
            seller_data['total_revenue'] = seller_data.get('total_revenue', 0)
            seller_data['total_profit'] = seller_data.get('total_profit', 0)
            seller_data['total_bonus_profit'] = seller_data.get('total_bonus_profit', 0)
            seller_data['total_non_liquid_revenue'] = seller_data.get('total_non_liquid_revenue', 0)
        
        # Выводим статистику по типам
        print(f"📊 СТАТИСТИКА ПО ТИПАМ ПРОДАЖ:")
        total_чеки = sum(s['продажи_чеки']['выручка'] for s in sales_data.values())
        total_опт = sum(s['продажи_опт']['прибыль'] for s in sales_data.values())
        total_прочая = sum(s['продажи_прочая']['прибыль'] for s in sales_data.values())
        
        print(f"   • Розничная (по чекам): {total_чеки:,.0f} руб.")
        print(f"   • Оптовая продажа: {total_опт:,.0f} руб.")
        print(f"   • Розничная прочая: {total_прочая:,.0f} руб.")
        
        return sales_data
        
    except Exception as e:
        import traceback
        print(f"❌ Ошибка парсинга: {str(e)}")
        traceback.print_exc()
        return {}
