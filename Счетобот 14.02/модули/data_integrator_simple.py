import pandas as pd
import numpy as np
from datetime import datetime

class DataIntegrator:
    @staticmethod
    def normalize_name(full_name):
        if pd.isna(full_name) or not isinstance(full_name, str):
            return ''
        return ' '.join(str(full_name).strip().split()).upper()
    
    @staticmethod
    def integrate_schedule_and_staff(schedule_data, staff_data):
        if 'error' in schedule_data or not staff_data.get('success'):
            print("❌ Нет данных для интеграции")
            return None
        
        schedule_df = schedule_data.get('employees_df', pd.DataFrame())
        schedule_dict = {}
        
        for _, row in schedule_df.iterrows():
            fio = row.get('ФИО', '')
            if fio:
                fio_norm = DataIntegrator.normalize_name(fio)
                schedule_dict[fio_norm] = {
                    'ФИО_оригинал': fio,
                    'Часы_всего': float(row.get('Часы_всего', 0)),
                    'Выходные_дни': int(row.get('Выходные_дни', 0)),
                    'Отпуск_дни': int(row.get('Отпуск_дни', 0)),
                    'Больничные_дни': int(row.get('Больничные_дни', 0))
                }
        
        print(f"📅 Данных из графика: {len(schedule_dict)} сотрудников")
        
        staff_employees = staff_data.get('employees', [])
        staff_dict = {}
        
        for emp in staff_employees:
            fio_norm = emp.get('ФИО_норм', '')
            dept_name = emp.get('Отдел', '')
            
            if not dept_name or dept_name == 'Не указан':
                continue
                
            if fio_norm:
                staff_dict[fio_norm] = {
                    'ФИО_оригинал': emp.get('ФИО', ''),
                    'Филиал': emp.get('Филиал', 'Не указан'),
                    'Отдел': dept_name,
                    'Директор_филиала': emp.get('Директор_филиала', 'Не указан')
                }
        
        print(f"👥 Данных о сотрудниках (с отделом): {len(staff_dict)} сотрудников")
        
        integrated_records = []
        all_employees = set(list(schedule_dict.keys()) + list(staff_dict.keys()))
        
        for fio_norm in all_employees:
            schedule_info = schedule_dict.get(fio_norm, {})
            staff_info = staff_dict.get(fio_norm, {})
            
            if fio_norm not in staff_dict:
                continue
                
            fio_original = schedule_info.get('ФИО_оригинал') or staff_info.get('ФИО_оригинал') or fio_norm
            
            record = {
                'ФИО': fio_original,
                'ФИО_норм': fio_norm,
                'Часы_всего': schedule_info.get('Часы_всего', 0.0),
                'Выходные_дни': schedule_info.get('Выходные_дни', 0),
                'Отпуск_дни': schedule_info.get('Отпуск_дни', 0),
                'Больничные_дни': schedule_info.get('Больничные_дни', 0),
                'Источник_график': 'Да' if fio_norm in schedule_dict else 'Нет',
                'Филиал': staff_info.get('Филиал', 'Не указан'),
                'Отдел': staff_info.get('Отдел', 'Не указан'),
                'Директор_филиала': staff_info.get('Директор_филиала', 'Не указан'),
                'Источник_сотрудники': 'Да',
                'Данные_продаж': None,
                'Выручка': 0.0,
                'Прибыль': 0.0,
                'Бонусные_продажи': 0.0,
                'Неликвидные_продажи': 0.0
            }
            
            integrated_records.append(record)
        
        if not integrated_records:
            print("❌ Нет данных для интеграции после фильтрации")
            return None
            
        integrated_df = pd.DataFrame(integrated_records)
        integrated_df = integrated_df.sort_values(['Филиал', 'Отдел', 'ФИО'])
        integrated_df = integrated_df[integrated_df['Отдел'] != 'Не указан']
        
        print(f"✅ Итоговый размер: {len(integrated_df)} сотрудников с отделами")
        
        no_schedule = integrated_df[integrated_df['Источник_график'] == 'Нет']
        if not no_schedule.empty:
            print(f"⚠️  Сотрудники без графика ({len(no_schedule)}):")
            for _, row in no_schedule.iterrows():
                print(f"   • {row['ФИО']} - {row['Отдел']}")

        return integrated_df
    
    @staticmethod
    def add_sales_data(integrated_df, sales_data, manager=None):
        if integrated_df.empty or not sales_data:
            return integrated_df
        
        # 1. СОЗДАЕМ НОВЫЙ DATAFRAME
        df = integrated_df.copy()
        df.index = range(len(df))
        
        # 2. ИНИЦИАЛИЗИРУЕМ КОЛОНКИ
        df['Данные_продаж'] = None
        df['Выручка'] = 0.0
        df['Прибыль'] = 0.0
        df['Бонусные_продажи'] = 0.0
        df['Неликвидные_продажи'] = 0.0
        df['Заказные_данные'] = None
        
        # 3. СОЗДАЕМ СЛОВАРЬ НОРМАЛИЗОВАННЫХ ИМЕН
        normalized_sales = {}
        for seller_name, sales_info in sales_data.items():
            seller_norm = DataIntegrator.normalize_name(seller_name)
            normalized_sales[seller_norm] = sales_info
        
        # 4. ЗАПОЛНЯЕМ ДАННЫМИ (ВЕКТОРИЗОВАННО)
        for norm_name, sales_info in normalized_sales.items():
            mask = df['ФИО_норм'] == norm_name
            
            if mask.any():
                df.loc[mask, 'Данные_продаж'] = df.loc[mask, 'ФИО_норм'].apply(lambda x: sales_info)
                df.loc[mask, 'Выручка'] = float(sales_info.get('total_revenue', 0))
                df.loc[mask, 'Прибыль'] = float(sales_info.get('total_profit', 0))
                df.loc[mask, 'Бонусные_продажи'] = float(sales_info.get('total_bonus_revenue', 0))
                df.loc[mask, 'Неликвидные_продажи'] = float(sales_info.get('total_non_liquid_revenue', 0))
        
        # 5. ЗАПОЛНЯЕМ ДАННЫЕ ПО ЗАКАЗНЫМ ТОВАРАМ
        if manager and hasattr(manager, 'zakaz_data') and manager.zakaz_data and manager.zakaz_data.get('success'):
            zakaz_dict = manager.zakaz_data.get('data', {})
            for norm_name, zakaz_info in zakaz_dict.items():
                mask = df['ФИО_норм'] == norm_name
                if mask.any():
                    df.loc[mask, 'Заказные_данные'] = df.loc[mask, 'ФИО_норм'].apply(lambda x: zakaz_info)
        
        print(f"💰 Добавлены данные о продажах для {(df['Выручка'] > 0).sum()} сотрудников")
        return df
    
    @staticmethod
    def add_urs_settings(integrated_df, urs_data):
        if integrated_df.empty or not urs_data.get('success'):
            return integrated_df
        
        # СОЗДАЕМ КОПИЮ и СБРАСЫВАЕМ ИНДЕКС - ЭТО РЕШЕНИЕ!
        integrated_df = integrated_df.copy().reset_index(drop=True)
        
        departments_settings = urs_data.get('departments', {})
        оклад_I2 = urs_data.get('оклад_I2', 0)
        
        print(f"🔍 Отделов в УРС: {len(departments_settings)}")
        print(f"💰 Оклад из ячейки I2: {оклад_I2:,.0f} руб.")
        
        for idx, row in integrated_df.iterrows():
            dept_name = row['Отдел']
            dept_settings = departments_settings.get(dept_name)
            
            if dept_settings:
                # БАЗОВЫЕ НАСТРОЙКИ
                integrated_df.loc[idx, 'Базовая_часть'] = float(dept_settings.get('базовая_часть', 0))
                integrated_df.loc[idx, 'Оклад'] = float(оклад_I2)
                integrated_df.loc[idx, 'Минималка_отдела'] = float(dept_settings.get('минималка', 0))
                integrated_df.loc[idx, 'Средняя_ЗП'] = float(dept_settings.get('средняя_зп', 0))
                integrated_df.loc[idx, 'Неликвиды_в_котле'] = dept_settings.get('неликвиды_в_котле', False)
                integrated_df.loc[idx, 'Неликвид_процент'] = float(dept_settings.get('неликвид_процент', 0.0))
                
                # КОЭФФИЦИЕНТЫ
                integrated_df.loc[idx, 'Коэф_обычных'] = float(dept_settings.get('коэф_обычных', 0.0))
                integrated_df.loc[idx, 'Коэф_бонусных'] = float(dept_settings.get('коэф_бонусных', 0.0))
                integrated_df.loc[idx, 'Коэф_неликвидов'] = float(dept_settings.get('коэф_неликвидов', 0.0))
                integrated_df.loc[idx, 'Коэф_оптовых'] = float(dept_settings.get('коэф_оптовых', 0.0))
                
                # ГАРАНТИИ (1-5 места)
                integrated_df.loc[idx, 'Гарантия_1'] = float(dept_settings.get('гарантия_1', 0.0))
                integrated_df.loc[idx, 'Гарантия_2'] = float(dept_settings.get('гарантия_2', 0.0))
                integrated_df.loc[idx, 'Гарантия_3'] = float(dept_settings.get('гарантия_3', 0.0))
                integrated_df.loc[idx, 'Гарантия_4'] = float(dept_settings.get('гарантия_4', 0.0))
                integrated_df.loc[idx, 'Гарантия_5'] = float(dept_settings.get('гарантия_5', 0.0))
                
                # НОРМЫ ЧАСОВ
                integrated_df.loc[idx, 'Тип_нормы'] = dept_settings.get('тип_нормы', '')
                integrated_df.loc[idx, 'Норма_часов_из_УРС'] = dept_settings.get('норма_часов', None)
            else:
                # НЕТ НАСТРОЕК - ЗАПОЛНЯЕМ НУЛЯМИ
                integrated_df.loc[idx, 'Базовая_часть'] = 0.0
                integrated_df.loc[idx, 'Оклад'] = 0.0
                integrated_df.loc[idx, 'Минималка_отдела'] = 0.0
                integrated_df.loc[idx, 'Средняя_ЗП'] = 0.0
                integrated_df.loc[idx, 'Неликвиды_в_котле'] = False
                integrated_df.loc[idx, 'Неликвид_процент'] = 0.0
                integrated_df.loc[idx, 'Коэф_обычных'] = 0.0
                integrated_df.loc[idx, 'Коэф_бонусных'] = 0.0
                integrated_df.loc[idx, 'Коэф_неликвидов'] = 0.0
                integrated_df.loc[idx, 'Коэф_оптовых'] = 0.0
                integrated_df.loc[idx, 'Гарантия_1'] = 0.0
                integrated_df.loc[idx, 'Гарантия_2'] = 0.0
                integrated_df.loc[idx, 'Гарантия_3'] = 0.0
                integrated_df.loc[idx, 'Гарантия_4'] = 0.0
                integrated_df.loc[idx, 'Гарантия_5'] = 0.0
                integrated_df.loc[idx, 'Тип_нормы'] = ''
                integrated_df.loc[idx, 'Норма_часов_из_УРС'] = None
        
        return integrated_df
    
    @staticmethod
    def create_integrated_dataframe(manager, office_norm_hours=168):
        print("\n" + "="*60)
        print("ИНТЕГРАЦИЯ ДАННЫХ")
        print("="*60)
        
        integrated_df = DataIntegrator.integrate_schedule_and_staff(
            manager.schedule_data, 
            manager.staff_data
        )
        
        if integrated_df is None:
            return None
        
        float_columns = ['Выручка', 'Прибыль', 'Бонусные_продажи', 'Неликвидные_продажи', 
                 'Базовая_часть', 'Оклад', 'Минималка_отдела', 'Средняя_ЗП', 'Часы_всего',
                 'Коэф_обычных', 'Коэф_бонусных', 'Коэф_неликвидов', 'Коэф_оптовых',
                 'Гарантия_1', 'Гарантия_2', 'Гарантия_3', 'Гарантия_4', 'Гарантия_5']
        
        for col in float_columns:
            if col not in integrated_df.columns:
                integrated_df[col] = 0.0
            else:
                integrated_df[col] = integrated_df[col].astype(float)
        
        if manager.sales_data:
            integrated_df = DataIntegrator.add_sales_data(integrated_df, manager.sales_data, manager)
        
        if manager.urs_data and manager.urs_data.get('success'):
            integrated_df = DataIntegrator.add_urs_settings(integrated_df, manager.urs_data)
                
        integrated_df = DataIntegrator.add_calculated_fields(integrated_df, manager, office_norm_hours)
        
        print(f"\n✅ Интеграция завершена!")
        print(f"📊 Итоговая таблица: {len(integrated_df)} записей")
        print(f"   С графиком: {(integrated_df['Источник_график'] == 'Да').sum()}")
        print(f"   Со структурой: {(integrated_df['Источник_сотрудники'] == 'Да').sum()}")
        print(f"   С продажами: {len([x for x in integrated_df['Данные_продаж'] if x is not None])}")
        
        return integrated_df

    @staticmethod
    def add_calculated_fields(df, manager=None, office_norm_hours=168):
        if df.empty:
            return df
        
        # 1. СОЗДАЕМ НОВЫЙ DATAFRAME
        df_result = df.copy()
        df_result.index = range(len(df_result))
        
        # 2. ПРОВЕРКА НОРМЫ МАГАЗИНА
        shop_norm_hours = None
        if manager and hasattr(manager, 'shop_norm_hours') and manager.shop_norm_hours:
            shop_norm_hours = manager.shop_norm_hours
        else:
            raise ValueError("❌ ОШИБКА: Норма часов для магазина не рассчитана!")
        
        # 3. РАСЧЕТ НОРМЫ ЧАСОВ (ВЕКТОРИЗОВАННО)
        df_result['Норма_часов'] = 0.0
        
        # Магазин
        mask_shop = df_result['Тип_нормы'] == 'магазин'
        df_result.loc[mask_shop, 'Норма_часов'] = shop_norm_hours
        
        # Офис
        mask_office = df_result['Тип_нормы'] == 'офис'
        df_result.loc[mask_office, 'Норма_часов'] = office_norm_hours
        
        # 4. ПРОВЕРКА НА НУЛЕВЫЕ НОРМЫ
        zero_norms = (df_result['Норма_часов'] == 0).sum()
        if zero_norms > 0:
            unknown_types = df_result[df_result['Норма_часов'] == 0]['Тип_нормы'].unique()
            error_msg = f"❌ ОШИБКА: {zero_norms} сотрудников имеют норму часов = 0\n"
            error_msg += f"   Неизвестные типы норм: {list(unknown_types)}"
            raise ValueError(error_msg)
        
        # 5. РАСЧЕТ ПРОИЗВОДНЫХ ПОЛЕЙ
        df_result['Процент_нормы'] = (df_result['Часы_всего'] / df_result['Норма_часов'] * 100).round(1)
        df_result['Статус_часов'] = df_result['Процент_нормы'].apply(lambda x: 'Выполнено' if x >= 100 else 'Не выполнено')
        df_result['Есть_продажи'] = df_result['Выручка'].apply(lambda x: 'Да' if x > 0 else 'Нет')
        
        # Процент бонусов (без деления на ноль)
        df_result['Процент_бонусов'] = 0.0
        mask_sales = df_result['Выручка'] > 0
        df_result.loc[mask_sales, 'Процент_бонусов'] = (df_result.loc[mask_sales, 'Бонусные_продажи'] / 
                                                         df_result.loc[mask_sales, 'Выручка'] * 100).round(1)
        
        shop_count = mask_shop.sum()
        office_count = mask_office.sum()
        
        print(f"📊 НОРМЫ ЧАСОВ УСТАНОВЛЕНЫ:")
        print(f"  ✅ Магазин: {shop_norm_hours} часов ({shop_count} сотрудников)")
        print(f"  ✅ Офис: {office_norm_hours} часов ({office_count} сотрудников)")
        print(f"  ✅ Всего: {len(df_result)} сотрудников, 0 ошибок")
        
        return df_result

def preview_integrated_data(df, max_rows=10):
    if df is None or df.empty:
        print("❌ Нет данных для просмотра")
        return
    
    print("\n👁 ПРЕВЬЮ ИНТЕГРИРОВАННЫХ ДАННЫХ:")
    print("-" * 120)
    
    display_cols = ['ФИО', 'Филиал', 'Отдел', 'Часы_всего', 'Выручка', 'Оклад_отдела', 'Минималка_отдела', 'Норма_часов', 'Тип_нормы']
    existing_cols = [col for col in display_cols if col in df.columns]
    
    if existing_cols:
        preview_df = df[existing_cols].head(max_rows)
        pd.set_option('display.width', 120)
        pd.set_option('display.max_columns', None)
        
        print(preview_df.to_string(index=False))
        
        print("\n📊 СТАТИСТИКА:")
        print(f"Всего записей: {len(df)}")
        print(f"С графиком: {(df['Источник_график'] == 'Да').sum()}")
        print(f"Со структурой: {(df['Источник_сотрудники'] == 'Да').sum()}")
        print(f"С продажами: {(df['Выручка'] > 0).sum()}")
        print(f"Общая выручка: {df['Выручка'].sum():,.0f} руб.")
        print(f"Общая прибыль: {df['Прибыль'].sum():,.0f} руб.")
    else:
        print("Нет данных для отображения")
