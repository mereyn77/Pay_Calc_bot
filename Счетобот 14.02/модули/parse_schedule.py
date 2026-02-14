import pandas as pd
from datetime import datetime
import re
import warnings
warnings.filterwarnings('ignore')

def parse_schedule(file_path):
    """
    Парсит график из Excel-файла.
    
    Возвращает:
    - period: строка периода
    - employees_df: DataFrame с колонками ['ФИО', 'Часы_всего', 'Выходные_дни', 'Отпуск_дни', 'Невыход_дни']
    """
    
    try:
        # Загружаем файл без заголовков
        df = pd.read_excel(file_path, header=None, dtype=str)
        
        # 1. Находим строку с периодом (ищем "С ... по ...")
        period = ""
        period_row = -1
        
        for i in range(min(10, len(df))):
            for j in range(df.shape[1]):
                cell = str(df.iloc[i, j])
                if 'с ' in cell.lower() and ' по ' in cell.lower():
                    period = cell.strip()
                    period_row = i
                    break
            if period:
                break
        
        # 2. Находим строку с заголовком "Сотрудник" или "ФИО"
        header_row = -1
        for i in range(len(df)):
            for j in range(df.shape[1]):
                cell = str(df.iloc[i, j]).lower()
                if 'сотрудник' in cell or 'фио' in cell:
                    header_row = i
                    break
            if header_row != -1:
                break
        
        if header_row == -1:
            return {"error": "Не найдена строка с заголовком 'Сотрудник' или 'ФИО'"}
        
        # 3. Находим колонку с итоговыми часами
        hours_col = -1
        for j in range(df.shape[1]):
            # Проверяем несколько строк после заголовка
            for check_row in range(header_row, min(header_row + 3, len(df))):
                cell = str(df.iloc[check_row, j]).lower()
                if 'итого' in cell and 'час' in cell:
                    hours_col = j
                    break
            if hours_col != -1:
                break
        
        if hours_col == -1:
            return {"error": "Не найдена колонка с итоговыми часами"}
        
        # 4. Собираем данные сотрудников
        employees_data = []
        
        # Начинаем со строки после заголовка
        for i in range(header_row + 1, len(df)):
            # Получаем ФИО (первая непустая ячейка в строке)
            fio = ""
            for j in range(df.shape[1]):
                cell = str(df.iloc[i, j]).strip()
                if cell and cell.lower() not in ['nan', 'none', '']:
                    # Проверяем, что это похоже на ФИО (не дата, не число часов)
                    if (len(cell.split()) >= 2 and  # хотя бы 2 слова
                        not any(c.isdigit() for c in cell[:5]) and  # не начинается с цифр
                        not cell.lower().startswith('итого')):
                        fio = cell
                        break
            
            if not fio:  # Если ФИО не найдено, пропускаем строку
                continue
            
            # Получаем часы
            try:
                hours_cell = str(df.iloc[i, hours_col])
                total_hours = float(hours_cell.replace(',', '.')) if hours_cell and hours_cell.lower() not in ['nan', 'none', ''] else 0.0
            except:
                total_hours = 0.0
            
            # Подсчитываем выходные, отпуск и невыходы (анализируем все ячейки строки)
            weekend_days = 0
            vacation_days = 0
            no_show_days = 0  # Невыходы
            sick_days = 0
            
            for j in range(df.shape[1]):
                if j == hours_col:  # Пропускаем колонку с часами
                    continue
                    
                cell = str(df.iloc[i, j]).strip().upper()
                if not cell or cell in ['NAN', 'NONE', '']:
                    continue
                
                # Учитываем различные варианты обозначений
                # ТОЛЬКО одна буква "Н" (не "РН", не "Н/Я" и т.д.)
                if cell == 'Н':
                    no_show_days += 1
                elif 'О' in cell and len(cell) <= 2:  # О, ОТ, ОТП
                    vacation_days += 1
                elif 'В' in cell and len(cell) <= 2:  # В, ВЫХ
                    weekend_days += 1
                elif 'Б' in cell and len(cell) <= 2:
                    sick_days += 1
            
            employees_data.append({
                'ФИО': fio,
                'Часы_всего': total_hours,
                'Выходные_дни': weekend_days,
                'Отпуск_дни': vacation_days,
                'Невыход_дни': no_show_days,
                'Больничные_дни': sick_days
            })
        
        # Создаем DataFrame
        if employees_data:
            employees_df = pd.DataFrame(employees_data)
        else:
            employees_df = pd.DataFrame(columns=['ФИО', 'Часы_всего', 'Выходные_дни', 'Отпуск_дни', 'Невыход_дни'])
        
        return {
            'period': period,
            'employees_df': employees_df,
            'period_row': period_row,
            'header_row': header_row,
            'hours_col': hours_col
        }
        
    except Exception as e:
        return {"error": f"Ошибка при чтении файла: {str(e)}"}

def print_schedule_results(result):
    """Выводит результаты парсинга в виде таблицы"""
    
    if 'error' in result:
        print(f"❌ Ошибка: {result['error']}")
        return
    
    print("=" * 70)
    print("РЕЗУЛЬТАТЫ ПАРСИНГА ГРАФИКА")
    print("=" * 70)
    
    print(f"📅 Период: {result['period']}")
    print(f"📊 Найдено сотрудников: {len(result['employees_df'])}")
    print(f"🔍 Строка заголовка: {result['header_row'] + 1}")
    print(f"🔍 Колонка с часами: {result['hours_col'] + 1}")
    
    print("\n" + "=" * 70)
    print("ТАБЛИЦА СОТРУДНИКОВ")
    print("=" * 70)
    
    if not result['employees_df'].empty:
        # Форматируем вывод
        df_display = result['employees_df'].copy()
        df_display['Часы_всего'] = df_display['Часы_всего'].round(1)
        
        # Показываем первые 20 строк
        pd.set_option('display.max_rows', 20)
        pd.set_option('display.width', 100)
        
        print(df_display.to_string(index=False))
        
        # Статистика
        print("\n" + "=" * 70)
        print("СТАТИСТИКА")
        print("=" * 70)
        
        total_hours = df_display['Часы_всего'].sum()
        avg_hours = df_display['Часы_всего'].mean()
        total_weekend = df_display['Выходные_дни'].sum()
        total_vacation = df_display['Отпуск_дни'].sum()
        total_no_show = df_display['Невыход_дни'].sum()
        
        print(f"Всего часов отработано: {total_hours:.1f}")
        print(f"Среднее часов на сотрудника: {avg_hours:.1f}")
        print(f"Всего выходных дней: {total_weekend}")
        print(f"Всего отпускных дней: {total_vacation}")
        print(f"Всего невыходов: {total_no_show}")
    else:
        print("Нет данных о сотрудниках")

# Пример использования:
if __name__ == "__main__":
    # Тестируем на вашем файле
    file_path = "График декабрь.xls"  # или полный путь к файлу
    
    print("🔍 Начинаю парсинг файла графика...")
    result = parse_schedule(file_path)
    
    print_schedule_results(result)
