# parse_bonus_integrated.py
import pandas as pd
import re

def parse_bonus_items_improved(file_path):
    """
    Улучшенный парсер бонусных позиций
    
    Возвращает словарь с:
    - bonus_items: set кодов бонусных товаров
    - non_liquid_items: set кодов неликвидов
    - items_info: dict с полной информацией {код: {статус, название, ...}}
    - statistics: статистика по файлу
    """
    
    print(f"🎁 Парсинг файла бонусов: {file_path}")
    
    try:
        # Читаем файл
        df = pd.read_excel(file_path, header=None, dtype=str)
        
        print(f"  📊 Размер файла: {len(df)} строк × {len(df.columns)} колонок")
        
        # 1. Находим колонки
        col_mapping = {}
        
        # Проходим по первым строкам для поиска заголовков
        for header_row in range(min(5, len(df))):
            for col_idx in range(len(df.columns)):
                cell = str(df.iloc[header_row, col_idx]).lower().strip()
                
                if 'код' in cell:
                    col_mapping['код'] = col_idx
                elif 'тмц' in cell or 'наимен' in cell or 'товар' in cell:
                    col_mapping['товар'] = col_idx
                elif 'статус' in cell or 'тип' in cell or 'бонус' in cell:
                    col_mapping['статус'] = col_idx
        
        # Если не нашли по заголовкам, используем логику из оригинала
        if not col_mapping:
            print("  ⚠️  Заголовки не найдены, использую стандартные позиции")
            if len(df.columns) >= 5:
                col_mapping = {'код': 0, 'статус': 4}
                if len(df.columns) > 1:
                    col_mapping['товар'] = 1
        
        print(f"  📋 Определены колонки: {col_mapping}")
        
        # 2. Определяем стартовую строку данных
        start_row = 0
        for i in range(min(10, len(df))):
            # Ищем строку, где в колонке кода есть числовое значение
            if 'код' in col_mapping:
                code_cell = str(df.iloc[i, col_mapping['код']])
                if re.match(r'^\d+$', code_cell.strip()):
                    start_row = i
                    break
        
        print(f"  🔍 Данные начинаются с строки: {start_row + 1}")
        
        # 3. Обрабатываем данные
        bonus_items = set()
        non_liquid_items = set()
        items_info = {}
        
        processed = 0
        skipped = 0
        
        for i in range(start_row, len(df)):
            # Получаем код
            code = ''
            if 'код' in col_mapping:
                code_raw = str(df.iloc[i, col_mapping['код']])
                code = code_raw.strip()
            
            # Пропускаем пустые
            if not code or code.lower() in ['nan', 'none', '']:
                skipped += 1
                continue
            
            # Пропускаем заголовки
            if 'код' in code.lower():
                continue
            
            # Проверяем формат кода (должен содержать цифры)
            if not re.search(r'\d', code):
                skipped += 1
                continue
            
            # Получаем статус
            status = ''
            if 'статус' in col_mapping:
                status_raw = str(df.iloc[i, col_mapping['статус']])
                # Нормализация: удаляем лишние пробелы, оставляем один пробел между словами
                status = ' '.join(status_raw.strip().split()).lower()
            
            # Получаем название товара
            товар = ''
            if 'товар' in col_mapping:
                товар_raw = str(df.iloc[i, col_mapping['товар']])
                товар = товар_raw.strip()
            
            # Определяем категорию
            if processed < 5:  # Отладка для первых 5 товаров
                print(f"    Товар {code}: статус='{status}', товар='{товар[:30]}...'")
            
            # Нормализуем статус еще раз для надежности
            status_normalized = ' '.join(status.lower().split())
            
            # Проверяем на неликвиды (приоритет выше, чем бонусы)
            is_non_liquid = False
            is_bonus = False
            
            # Варианты для неликвидов: "бонус уценка", "уценка бонус" (в любом порядке с пробелами)
            if ('бонус' in status_normalized and 'уценка' in status_normalized):
                is_non_liquid = True
            # Чистые бонусы (без слова "уценка")
            elif 'бонус' in status_normalized and 'уценка' not in status_normalized:
                is_bonus = True
            
            if is_bonus:
                bonus_items.add(code)
                items_info[code] = {
                    'статус': 'бонус',
                    'название': товар,
                    'строка': i + 1,
                    'исходный_статус': status_raw if 'статус' in col_mapping else ''
                }
            elif is_non_liquid:
                non_liquid_items.add(code)
                items_info[code] = {
                    'статус': 'неликвид',
                    'название': товар,
                    'строка': i + 1,
                    'исходный_статус': status_raw if 'статус' in col_mapping else ''
                }
            else:
                # Товары без статуса бонуса
                items_info[code] = {
                    'статус': 'обычный',
                    'название': товар,
                    'строка': i + 1,
                    'исходный_статус': status_raw if 'статус' in col_mapping else ''
                }
            
            processed += 1
        
        # 4. Формируем результат
        result = {
            'success': True,
            'bonus_items': bonus_items,
            'non_liquid_items': non_liquid_items,
            'items_info': items_info,
            'statistics': {
                'total_processed': processed,
                'total_skipped': skipped,
                'bonus_count': len(bonus_items),
                'non_liquid_count': len(non_liquid_items),
                'total_unique': len(items_info),
                'columns_mapped': col_mapping,
                'start_row': start_row
            }
        }
        
        print(f"\n  ✅ Обработано: {processed} товаров")
        print(f"  ✗ Пропущено: {skipped} строк")
        print(f"  🎁 Бонусных: {len(bonus_items)}")
        print(f"  📦 Неликвидов: {len(non_liquid_items)}")
        
        # Примеры для проверки
        if bonus_items:
            print(f"\n  📋 Примеры бонусных товаров (первые 3):")
            for i, code in enumerate(list(bonus_items)[:3]):
                info = items_info.get(code, {})
                name = info.get('название', 'Нет названия')[:40]
                status_info = info.get('статус', '?')
                print(f"    {i+1}. {code} - {name}... [{status_info}]")
        
        if non_liquid_items:
            print(f"\n  📋 Примеры неликвидов (первые 3):")
            for i, code in enumerate(list(non_liquid_items)[:3]):
                info = items_info.get(code, {})
                name = info.get('название', 'Нет названия')[:40]
                original_status = info.get('исходный_статус', '')[:20]
                print(f"    {i+1}. {code} - {name}... ['{original_status}']")
        
        return result
        
    except Exception as e:
        import traceback
        print(f"❌ Ошибка парсинга: {str(e)}")
        return {
            'success': False,
            'error': f"Ошибка парсинга: {str(e)}",
            'traceback': traceback.format_exc()
        }

def check_item_status(item_code, bonus_data):
    """
    Проверяет статус товара по коду
    
    Возвращает:
    - 'бонус' - бонусный товар
    - 'неликвид' - неликвид
    - 'обычный' - обычный товар
    - None - код не найден
    """
    if not bonus_data.get('success'):
        return None
    
    item_info = bonus_data.get('items_info', {}).get(str(item_code))
    if item_info:
        return item_info.get('статус')
    return None

# Интеграция с DataManager
class BonusDataProcessor:
    """Адаптер для интеграции с DataManager"""
    
    @staticmethod
    def process_for_datamanager(bonus_result):
        """
        Преобразует результат парсера для DataManager
        
        Возвращает:
        - bonus_codes: set бонусных кодов
        - non_liquid_codes: set неликвидов
        - items_dict: полный словарь товаров
        """
        if not bonus_result.get('success'):
            return set(), set(), {}
        
        return (
            bonus_result['bonus_items'],
            bonus_result['non_liquid_items'],
            bonus_result['items_info']
        )

# Пример использования
if __name__ == "__main__":
    # Тестируем парсер
    file_path = "Список бонусные позиции Декабрь.xlsx"
    result = parse_bonus_items_improved(file_path)
    
    if result['success']:
        print("\n" + "="*60)
        print("ДЕТАЛЬНАЯ СТАТИСТИКА")
        print("="*60)
        
        stats = result['statistics']
        print(f"Всего уникальных товаров: {stats['total_unique']}")
        print(f"Бонусных товаров: {stats['bonus_count']}")
        print(f"Неликвидов: {stats['non_liquid_count']}")
        
        # Пример проверки статуса
        if result['bonus_items']:
            test_code = list(result['bonus_items'])[0]
            status = check_item_status(test_code, result)
            print(f"\nПроверка статуса товара {test_code}: {status}")
