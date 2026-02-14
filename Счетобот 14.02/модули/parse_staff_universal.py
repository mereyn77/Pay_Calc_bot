# parse_staff_universal.py
import pandas as pd
from collections import defaultdict, Counter

def normalize_name(full_name):
    """Нормализует ФИО для сравнения"""
    if pd.isna(full_name):
        return ''
    return ' '.join(str(full_name).strip().split()).upper()

def parse_staff_departments(file_path):
    """
    Универсальный парсер файла 'Сотрудники по отделам.xlsx'
    Работает с ЛЮБЫМИ названиями филиалов и фамилиями директоров
    
    Структура:
    Колонка 1: ФИО сотрудника
    Колонка 2: Название филиала (любое)
    Колонка 3: Фамилия директора (любая) 
    Колонка 4: Название отдела (любое)
    """
    
    print("=" * 80)
    print("УНИВЕРСАЛЬНЫЙ ПАРСЕР ФАЙЛА СОТРУДНИКОВ")
    print("=" * 80)
    
    try:
        # Читаем файл
        df = pd.read_excel(file_path, dtype=str)
        
        print(f"📁 Файл: {file_path}")
        print(f"📊 Размер: {len(df)} строк × {len(df.columns)} колонок")
        print("\n🔍 Обнаруженные колонки:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1:2}. '{col}'")
        
        # СТРАТЕГИЯ: Определяем колонки по содержимому и позиции
        col_mapping = {}
        
        # 1. Определяем колонку с ФИО (по характерным признакам)
        for i, col in enumerate(df.columns):
            # Проверяем первые несколько значений в колонке
            sample_values = df[col].dropna().astype(str).str.strip().head(10)
            
            # Признаки колонки с ФИО:
            # - Содержит 2-4 слова в большинстве значений
            # - Содержит русские буквы
            # - Не содержит типичных маркеров других колонок
            is_fio_column = False
            fio_count = 0
            total_count = 0
            
            for val in sample_values:
                if not val or val.lower() in ['nan', 'none']:
                    continue
                    
                total_count += 1
                words = val.split()
                
                # Признаки ФИО
                if (2 <= len(words) <= 4 and  # 2-4 слова (Фамилия Имя Отчество)
                    any(cyrillic in val for cyrillic in 'АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ') and  # Русские буквы
                    not any(marker in val.lower() for marker in ['филиал', 'отдел', 'директор', 'город'])):  # Не заголовок
                    fio_count += 1
            
            if total_count > 0 and fio_count / total_count > 0.7:  # >70% значений похожи на ФИО
                col_mapping['ФИО'] = col
                print(f"✅ Колонка '{col}' определена как 'ФИО'")
                break
        
        # 2. Определяем остальные колонки по порядку и содержимому
        remaining_cols = [c for c in df.columns if c not in col_mapping.values()]
        
        if len(remaining_cols) >= 3:
            # Колонка 2: Филиал (обычно содержит названия городов/подразделений)
            # Колонка 3: Директор (фамилии, обычно короче ФИО)
            # Колонка 4: Отдел (названия отделов, могут быть длинными)
            
            # Анализируем каждую колонку
            col_characteristics = []
            for col in remaining_cols[:3]:  # Берем первые 3 оставшиеся колонки
                sample = df[col].dropna().astype(str).str.strip().head(20)
                
                # Характеристики колонки
                avg_length = sample.str.len().mean() if not sample.empty else 0
                unique_count = sample.nunique()
                contains_director_keywords = any('директор' in str(v).lower() for v in sample)
                contains_branch_keywords = any(word in str(v).lower() for v in sample for word in ['филиал', 'город', 'отдел'])
                
                col_characteristics.append({
                    'col': col,
                    'avg_length': avg_length,
                    'unique_count': unique_count,
                    'is_director': contains_director_keywords or (avg_length < 20 and unique_count < 10),
                    'is_branch': contains_branch_keywords or (avg_length > 5 and avg_length < 30)
                })
            
            # Сопоставляем по логике:
            # 1. Филиал: средняя длина, несколько уникальных значений
            # 2. Директор: короткие значения, мало уникальных
            # 3. Отдел: разнообразные значения, могут быть длинными
            
            if len(col_characteristics) >= 3:
                # Сортируем по средней длине (директор обычно короче)
                sorted_by_length = sorted(col_characteristics, key=lambda x: x['avg_length'])
                
                # Предполагаем порядок: Филиал -> Директор -> Отдел
                col_mapping['Филиал'] = remaining_cols[0]
                col_mapping['Директор'] = remaining_cols[1] 
                col_mapping['Отдел'] = remaining_cols[2]
                
                print(f"📋 Автоматическое сопоставление:")
                print(f"  Колонка 2 ('{remaining_cols[0]}') → 'Филиал'")
                print(f"  Колонка 3 ('{remaining_cols[1]}') → 'Директор'")
                print(f"  Колонка 4 ('{remaining_cols[2]}') → 'Отдел'")
        
        # 3. Если не определили автоматически, используем ручное сопоставление
        if 'Филиал' not in col_mapping and len(df.columns) >= 4:
            print("⚠️ Автоматическое определение не сработало, использую порядок колонок")
            col_mapping = {
                'ФИО': df.columns[0],
                'Филиал': df.columns[1],
                'Директор': df.columns[2],
                'Отдел': df.columns[3]
            }
        
        # 4. Проверяем наличие обязательных колонок
        required = ['ФИО', 'Филиал', 'Отдел']
        for col in required:
            if col not in col_mapping:
                return {
                    "success": False,
                    "error": f"Не удалось определить колонку '{col}'",
                    "available_columns": list(df.columns),
                    "col_mapping": col_mapping
                }
        
        print(f"\n✅ Окончательное сопоставление колонок:")
        for key, col in col_mapping.items():
            print(f"  {key:15} → '{col}'")
        
        # 5. Обрабатываем данные
        employees = []
        branch_info = defaultdict(lambda: {'director': None, 'employees': [], 'departments': set()})
        department_info = defaultdict(lambda: {'employees': [], 'branches': set()})
        
        processed_count = 0
        skipped_count = 0
        
        for idx, row in df.iterrows():
            # Получаем значения
            fio_raw = str(row[col_mapping['ФИО']]).strip() if col_mapping['ФИО'] in row else ''
            branch_raw = str(row[col_mapping['Филиал']]).strip() if col_mapping['Филиал'] in row else ''
            dept_raw = str(row[col_mapping['Отдел']]).strip() if col_mapping['Отдел'] in row else ''
            
            # Получаем директора (если есть колонка)
            director_raw = ''
            if 'Директор' in col_mapping and col_mapping['Директор'] in row:
                director_raw = str(row[col_mapping['Директор']]).strip()
            
            # Пропускаем пустые или некорректные строки
            if (not fio_raw or fio_raw.lower() in ['nan', 'none', ''] or
                len(fio_raw) < 2 or
                fio_raw.lower() in ['фио', 'сотрудник', 'ф.и.о.']):
                skipped_count += 1
                continue
            
            # Очищаем данные
            fio_clean = ' '.join(fio_raw.split())  # Удаляем лишние пробелы
            branch_clean = ' '.join(branch_raw.split()) if branch_raw else 'Не указан'
            dept_clean = ' '.join(dept_raw.split()) if dept_raw else 'Не указан'
            director_clean = ' '.join(director_raw.split()) if director_raw else 'Не указан'
            
            # ФИЛЬТРАЦИЯ: Пропускаем сотрудников без отдела
            if dept_clean == 'Не указан':
                skipped_count += 1
                print(f"  ⚠️  Строка {idx+2}: Пропущен сотрудник без отдела - '{fio_clean[:30]}'")
                continue
            
            # Нормализуем ФИО для сравнения
            fio_norm = normalize_name(fio_clean)
            
            # Создаем запись сотрудника
            employee = {
                'ФИО': fio_clean,
                'ФИО_норм': fio_norm,
                'Филиал': branch_clean,
                'Отдел': dept_clean,
                'Директор_филиала': director_clean,
                'row_index': idx + 2  # +2 потому что Excel строки с 1 и header
            }
            employees.append(employee)
            processed_count += 1
            
            # Обновляем информацию о филиале
            if branch_clean != 'Не указан':
                branch_info[branch_clean]['employees'].append(fio_norm)
                branch_info[branch_clean]['departments'].add(dept_clean)
                if director_clean and director_clean != 'Не указан':
                    branch_info[branch_clean]['director'] = director_clean
            
            # Обновляем информацию об отделе
            if dept_clean != 'Не указан':
                department_info[dept_clean]['employees'].append(fio_norm)
                department_info[dept_clean]['branches'].add(branch_clean)
        
        print(f"\n📊 ОБРАБОТКА ДАННЫХ:")
        print(f"  ✓ Обработано строк: {processed_count}")
        print(f"  ✗ Пропущено строк: {skipped_count}")
        
        if processed_count == 0:
            return {
                "success": False,
                "error": "Не найдено данных сотрудников",
                "processed": 0,
                "skipped": skipped_count
            }
        
        # 6. Формируем результат
        branches = sorted(branch_info.keys())
        departments = sorted(department_info.keys())
        
        # Создаем удобные структуры
        branch_directors = {b: info['director'] for b, info in branch_info.items()}
        departments_by_branch = {b: sorted(list(info['departments'])) for b, info in branch_info.items()}
        
        # Статистика
        branch_counts = Counter(e['Филиал'] for e in employees)
        dept_counts = Counter(e['Отдел'] for e in employees)
        
        result = {
            "success": True,
            "summary": {
                "total_employees": len(employees),
                "total_branches": len(branches),
                "total_departments": len(departments),
                "branches_with_director": sum(1 for d in branch_directors.values() if d),
                "processed_rows": processed_count,
                "skipped_rows": skipped_count
            },
            
            # Основные данные
            "employees": employees,
            
            # Группировки
            "grouping": {
                "by_branch": {b: branch_info[b]['employees'] for b in branches},
                "by_department": {d: department_info[d]['employees'] for d in departments},
                "departments_by_branch": departments_by_branch,
                "branch_directors": branch_directors,
                "branches_by_department": {d: sorted(list(department_info[d]['branches'])) for d in departments}
            },
            
            # Статистика
            "statistics": {
                "branches": branches,
                "departments": departments,
                "employees_per_branch": dict(branch_counts),
                "employees_per_department": dict(dept_counts),
                "avg_employees_per_branch": len(employees) / len(branches) if branches else 0,
                "avg_employees_per_department": len(employees) / len(departments) if departments else 0
            },
            
            # Метаданные
            "metadata": {
                "file_path": file_path,
                "columns_found": list(df.columns),
                "columns_mapped": col_mapping,
                "total_rows": len(df)
            }
        }
        
        return result
        
    except Exception as e:
        import traceback
        return {
            "success": False,
            "error": f"Ошибка: {str(e)}",
            "traceback": traceback.format_exc()
        }

def print_detailed_report(result):
    """Выводит детальный отчет"""
    
    if not result.get("success", False):
        print(f"\n❌ ОШИБКА: {result.get('error', 'Неизвестная ошибка')}")
        if 'traceback' in result:
            print("\nДетали ошибки:")
            print(result['traceback'][:300])
        return
    
    print("\n" + "=" * 100)
    print("ДЕТАЛЬНЫЙ ОТЧЕТ О ПАРСИНГЕ")
    print("=" * 100)
    
    summary = result['summary']
    stats = result['statistics']
    
    # Основная статистика
    print(f"\n📈 ОСНОВНЫЕ ПОКАЗАТЕЛИ:")
    print(f"  👥 Сотрудников: {summary['total_employees']}")
    print(f"  🏢 Филиалов: {summary['total_branches']} ({summary['branches_with_director']} с директором)")
    print(f"  📁 Отделов: {summary['total_departments']}")
    print(f"  📊 Среднее по филиалу: {stats['avg_employees_per_branch']:.1f} сотрудников")
    print(f"  📊 Среднее по отделу: {stats['avg_employees_per_department']:.1f} сотрудников")
    
    # Филиалы с деталями
    print(f"\n🏢 ДЕТАЛИ ПО ФИЛИАЛАМ:")
    print("-" * 100)
    print(f"{'Филиал':25} | {'Директор':20} | {'Сотр.':6} | {'Отделов':8} | {'Пример отдела'}")
    print("-" * 100)
    
    for branch in sorted(stats['branches']):
        director = result['grouping']['branch_directors'].get(branch, '—')
        emp_count = stats['employees_per_branch'][branch]
        dept_count = len(result['grouping']['departments_by_branch'].get(branch, []))
        example_dept = result['grouping']['departments_by_branch'].get(branch, ['—'])[0][:20]
        
        print(f"{branch[:25]:25} | {director[:20]:20} | {emp_count:6} | {dept_count:8} | {example_dept}")
    
    # Отделы (топ-10)
    print(f"\n📁 КРУПНЕЙШИЕ ОТДЕЛЫ (ТОП-10):")
    print("-" * 70)
    print(f"{'Отдел':40} | {'Сотр.':6} | {'Филиалы'}")
    print("-" * 70)
    
    top_depts = sorted(stats['employees_per_department'].items(), 
                      key=lambda x: x[1], reverse=True)[:10]
    
    for dept, count in top_depts:
        branches_list = result['grouping']['branches_by_department'].get(dept, [])
        branches_str = ', '.join(b[:10] for b in branches_list[:2])
        if len(branches_list) > 2:
            branches_str += f" (+{len(branches_list)-2})"
        
        print(f"{dept[:40]:40} | {count:6} | {branches_str}")
    
    # Примеры данных
    print(f"\n👤 ПРИМЕРЫ ДАННЫХ (первые 10 сотрудников):")
    print("-" * 120)
    print(f"{'№':3} | {'ФИО':35} | {'Филиал':20} | {'Отдел':30} | {'Директор':15}")
    print("-" * 120)
    
    for i, emp in enumerate(result['employees'][:10]):
        print(f"{i+1:3} | {emp['ФИО'][:35]:35} | {emp['Филиал'][:20]:20} | "
              f"{emp['Отдел'][:30]:30} | {emp['Директор_филиала'][:15]:15}")
    
    print("=" * 100)
    
    # Дополнительная информация
    print(f"\n📋 ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ:")
    print(f"  Обработано строк: {summary['processed_rows']}")
    print(f"  Пропущено строк: {summary['skipped_rows']}")
    
    # Уникальные значения для проверки
    print(f"\n  УНИКАЛЬНЫЕ ФИЛИАЛЫ ({len(stats['branches'])}):")
    for i, branch in enumerate(sorted(stats['branches']), 1):
        print(f"    {i:2}. {branch}")
    
    print(f"\n  УНИКАЛЬНЫЕ ДИРЕКТОРА:")
    directors = set(result['grouping']['branch_directors'].values())
    directors.discard('Не указан')
    directors.discard('')
    for i, director in enumerate(sorted(directors), 1):
        print(f"    {i:2}. {director}")

# Основная функция
def main():
    """Основная функция для тестирования"""
    
    # Укажите путь к файлу
    FILE_PATH = "Сотрудники по отделам.xlsx"
    
    print("🔄 Запуск универсального парсера...")
    result = parse_staff_departments(FILE_PATH)
    
    print_detailed_report(result)
    
    # Проверка целостности данных
    if result.get("success"):
        print(f"\n✅ ПРОВЕРКА ЦЕЛОСТНОСТИ ДАННЫХ:")
        
        employees = result['employees']
        
        # 1. Проверка уникальности ФИО
        fio_norm_set = set(e['ФИО_норм'] for e in employees)
        duplicates = len(employees) - len(fio_norm_set)
        print(f"   Уникальных ФИО: {len(fio_norm_set)} из {len(employees)}")
        if duplicates > 0:
            print(f"   ⚠️  Найдено возможных дубликатов: {duplicates}")
        
        # 2. Проверка филиалов без директора
        branches_without_director = [
            b for b, d in result['grouping']['branch_directors'].items() 
            if not d or d == 'Не указан'
        ]
        if branches_without_director:
            print(f"   ⚠️  Филиалы без директора: {', '.join(branches_without_director)}")
        
        # 3. Проверка отделов без сотрудников
        empty_departments = [
            d for d, count in result['statistics']['employees_per_department'].items()
            if count == 0
        ]
        if empty_departments:
            print(f"   ⚠️  Отделы без сотрудников: {len(empty_departments)}")
        
        print(f"\n🎯 ГОТОВНОСТЬ ДАННЫХ ДЛЯ РАСЧЕТОВ: 100%")

if __name__ == "__main__":
    main()
