import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
from utils import format_russian_number
import sys
from data_manager import DataManager
from модули.parse_zakaz_sales import parse_zakaz_sales

# Отладочный код
print(f"\n=== DEBUG: Проверка импортов ===")
print(f"Текущая директория: {os.getcwd()}")
print(f"Папка модули существует: {os.path.exists('модули')}")

# Импортируем DataManager
from data_manager import DataManager

class SalaryCalculatorApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Расчет зарплаты")
        self.root.geometry("990x700")
        
        # Создаем менеджер данных
        self.manager = DataManager()
        
        # Настройки
        self.office_norm_hours = 168  # значение по умолчанию
        
        # Создаем интерфейс
        self.create_widgets()
        
        # Показываем первое сообщение
        self.show_welcome_message()

        self.style = ttk.Style()
        self.style.configure('Active.TButton', 
                           font=('Segoe UI', 9),
                           foreground='blue')
        self.active_button = None
        
    def set_active_button(self, button_text):
        """Устанавливает активную кнопку"""
        # Сбрасываем предыдущую активную кнопку
        if self.active_button is not None:
            self.active_button.configure(style='TButton')
        
        # Находим все кнопки в левой панели (btn_frame)
        btn_frame = None
        for child in self.root.winfo_children():
            if isinstance(child, ttk.LabelFrame) and child['text'] == 'Действия':
                btn_frame = child
                break
        
        if btn_frame:
            for widget in btn_frame.winfo_children():
                if isinstance(widget, ttk.Button) and widget['text'] == button_text:
                    widget.configure(style='Active.TButton')
                    self.active_button = widget
                    return
        
    def show_welcome_message(self):
        """Показывает приветственное сообщение при запуске"""
        self.clear_log()
        self.log_message("👋 Вас приветствует система Счетобот (v.1.0 2026 (C))")
        self.log_message("="*60)
        self.log_message("📋 Инструкция:")
        self.log_message("")
        self.log_message("1. Сохраните следующие файлы в папке 'данные' в формате .xlsx:")
        self.log_message("   • УРС.xlsx (настройки расчета)")
        self.log_message("   • Список бонусные позиции.xlsx")
        self.log_message("   • График.xlsx (график работы)")
        self.log_message("   • Сотрудники по отделам.xlsx")
        self.log_message("   • Анализ продаж.xlsx")
        self.log_message("")
        self.log_message("Анализ продаж снимается со следующими настройками:")
        self.log_message("   • Отчетный период")
        self.log_message("   • Продавец")
        self.log_message("   • Вид продаж")
        self.log_message("   • Номенклатура")
        self.log_message("")
        self.log_message("2. Убедитесь, что все обрабатываемые файлы закрыты.")
        self.log_message("3. Последовательно нажимайте кнопки в левой панели.")
        self.log_message("="*60)
        self.update_status("Готов к работе. Нажмите '🔍 Проверить файлы' для начала")

    # Русский формат отображения чисел
    def _format_russian_number(self, num, decimal_places=0):
        """
        Форматирует число в русском стиле:
        - Тысячи разделяются пробелом
        - Дробная часть отделяется запятой
        """
        if num is None:
            return "0"
        
        try:
            # Для целых чисел
            if decimal_places == 0:
                num_int = int(round(float(num)))
                formatted = f"{abs(num_int):,}".replace(",", " ")
                return f"-{formatted}" if num_int < 0 else formatted
            # Для дробных чисел
            else:
                num_float = float(num)
                formatted = f"{abs(num_float):,.{decimal_places}f}".replace(",", " ").replace(".", ",")
                return f"-{formatted}" if num_float < 0 else formatted
        except (ValueError, TypeError):
            return str(num)
    
    def create_widgets(self):
        # Верхняя панель
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        ttk.Label(top_frame, text="📊 РАСЧЕТ ЗАРПЛАТЫ", font=('Arial', 18, 'bold')).pack()
        
        # Информация о периоде
        self.period_label = ttk.Label(top_frame, text="Период: не определен", font=('Arial', 10))
        self.period_label.pack(pady=5)

        # Настройки нормы часов
        settings_frame = ttk.Frame(top_frame)
        settings_frame.pack(pady=5)
        
        ttk.Label(settings_frame, text="Норма часов 'Офис':").pack(side=tk.LEFT, padx=5)
        self.office_norm_entry = ttk.Entry(settings_frame, width=10)
        self.office_norm_entry.insert(0, "168")
        self.office_norm_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(settings_frame, text="Применить", 
                  command=self.update_office_norm).pack(side=tk.LEFT, padx=5)
        
        # Кнопки действий слева
        btn_frame = ttk.LabelFrame(self.root, text="Действия", padding="10")
        btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        
        buttons = [
            ("🔍 Проверить файлы", self.check_files),
            ("📅 Определить период", self.detect_period),
            ("📊 Загрузить данные", self.load_data),
            ("🔄 Интегрировать данные", self.integrate_data),
            ("🧮 Рассчитать зарплату", self.calculate_salary),
            ("👁 Просмотреть данные", self.show_data_preview),
            ("📄 Создать отчет Excel", self.create_report),
            ("📋 Простой отчет", self.create_simple_report),
            ("📈 Дашборд", self.show_dashboard),
            ("💾 Сохранить результаты", self.save_results)
        ]
        
        for i, (text, command) in enumerate(buttons):
            btn = ttk.Button(btn_frame, text=text, command=command, width=25)
            btn.grid(row=i, column=0, padx=5, pady=5, sticky=tk.W)
        
        # Правая часть - информация и превью
        info_frame = ttk.LabelFrame(self.root, text="Информация и результаты", padding="10")
        info_frame.grid(row=1, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        
        # Текстовое поле для вывода
        self.info_text = tk.Text(info_frame, height=25, width=90, wrap=tk.WORD)
        self.info_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(info_frame, orient="vertical", command=self.info_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.info_text.configure(yscrollcommand=scrollbar.set)
        
        # Кнопки очистки и копирования
        text_btn_frame = ttk.Frame(info_frame)
        text_btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(text_btn_frame, text="Очистить", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(text_btn_frame, text="Копировать", command=self.copy_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(text_btn_frame, text="Сохранить в файл", command=self.save_log).pack(side=tk.LEFT, padx=5)
        
        # Статус бар внизу
        self.status_bar = ttk.Label(self.root, text="Готов к работе. Проверьте наличие файлов в папке 'данные'", 
                                    relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=10, pady=5)
    
    def log_message(self, message, color=None):
        """Добавляет цветное сообщение в информационную область"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Настраиваем теги для цветов (делаем один раз)
        if not hasattr(self, '_tags_configured'):
            self.info_text.tag_config("red", foreground="red")
            self.info_text.tag_config("green", foreground="#006400")  # Темно-зеленый
            self.info_text.tag_config("orange", foreground="#FF8C00")  # Темно-оранжевый
            self.info_text.tag_config("blue", foreground="#00008B")    # Темно-синий
            self.info_text.tag_config("purple", foreground="#4B0082")  # Индиго
            self.info_text.tag_config("gray", foreground="#696969")    # Темно-серый
            self.info_text.tag_config("black", foreground="black")
            self._tags_configured = True
        
        # Определяем цвет если не задан
        if color is None:
            if "❌" in message or "Ошибка" in message or "ошибка" in message.lower():
                color = "red"
            elif "✅" in message or "Успешно" in message or "готов" in message.lower():
                color = "green"
            elif "⚠️" in message or "Внимание" in message or "требуется" in message.lower():
                color = "orange"
            elif "💡" in message or "Подсказка" in message or "инструкция" in message.lower():
                color = "blue"
            elif "📅" in message or "📊" in message or "📋" in message:
                color = "purple"
            else:
                color = "black"
        
        # Вставляем сообщение
        self.info_text.insert(tk.END, f"[{timestamp}] ", "gray")
        self.info_text.insert(tk.END, message + "\n", color)
        self.info_text.see(tk.END)
        self.update_status(f"Выполнено: {message[:50]}...")
        self.root.update()
    
    def clear_log(self):
        """Очищает информационную область"""
        self.info_text.delete(1.0, tk.END)
    
    def copy_log(self):
        """Копирует логи в буфер обмена"""
        self.root.clipboard_clear()
        self.root.clipboard_append(self.info_text.get(1.0, tk.END))
        messagebox.showinfo("Копирование", "Логи скопированы в буфер обмена")
    
    def save_log(self):
        """Сохраняет логи в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(self.info_text.get(1.0, tk.END))
            self.log_message(f"Логи сохранены в файл: {filename}")
    
    def update_status(self, message):
        """Обновляет статус бар"""
        self.status_bar.config(text=message)
    
    def check_files(self):
        """Проверяет наличие файлов"""
        self.set_active_button("🔍 Проверить файлы")
        self.clear_log()
        self.log_message("🔍 Проверяю наличие файлов в папке 'данные'...")
        
        # Проверяем папку
        if not os.path.exists(self.manager.data_folder):
            self.log_message(f"❌ Папка '{self.manager.data_folder}' не найдена!")
            self.log_message("Создайте папку 'данные' и поместите в нее файлы:")
            self.log_message("  • УРС.xlsx (настройки расчета)")
            self.log_message("  • Список бонусные позиции.xlsx")
            self.log_message("  • График.xls (график работы)")
            self.log_message("  • Сотрудники по отделам.xlsx")
            self.log_message("  • Анализ продаж.xlsx")
            return False
        
        # Ищем файлы
        if self.manager.find_files_by_patterns():
            self.log_message("✅ Файлы найдены:")
            for file_type, filename in self.manager.found_files.items():
                filepath = os.path.join(self.manager.data_folder, filename)
                if os.path.exists(filepath):
                    size = os.path.getsize(filepath) / 1024
                    self.log_message(f"   • {file_type}: {filename} ({size:.1f} КБ)")
                else:
                    self.log_message(f"   • {file_type}: {filename} (файл не найден!)")
            
            # Проверяем наличие всех файлов
            required = ['urs', 'bonus', 'schedule', 'staff', 'sales']
            missing = [r for r in required if r not in self.manager.found_files]
            
            if missing:
                self.log_message(f"⚠️  Не найдены файлы: {', '.join(missing)}")
                return False
            else:
                self.log_message("✅ Все необходимые файлы найдены.")
                self.log_message("✅ Убедитесь, что файлы соответствуют отчетному периоду.")
                return True
        else:
            self.log_message("❌ Файлы не найдены!")
            return False

        

    def _calculate_shop_norm(self, period_str):
        """Рассчитывает норму часов для магазина по формуле"""
        import re
        from datetime import datetime
        import math
        
        if not period_str:
            self.log_message("   [Ошибка] Пустой период")
            return None
        
        try:
            # Извлекаем даты из строки периода
            # Поддерживаем разные форматы: 01.12.25, 01.12.2025
            dates = re.findall(r'\d{1,2}\.\d{1,2}\.\d{2,4}', period_str)
            
            if len(dates) >= 2:
                date1, date2 = dates[0], dates[1]
                
                # Пробуем разные форматы дат
                for date_format in ['%d.%m.%y', '%d.%m.%Y', '%d/%m/%y', '%d/%m/%Y']:
                    try:
                        start = datetime.strptime(date1, date_format)
                        end = datetime.strptime(date2, date_format)
                        break
                    except ValueError:
                        continue
                else:
                    self.log_message(f"   [Ошибка] Неподдерживаемый формат: '{date1}', '{date2}'")
                    return None
                
                days_in_month = (end - start).days + 1
                
                # Формула: (дней / 7) × 5 × 8
                shop_norm = days_in_month / 7 * 5 * 8
                # Округление вниз до 1 знака
                shop_norm = math.floor(shop_norm * 10) / 10
                
                self.log_message(f"   [Расчет] {days_in_month} дней / 7 × 5 × 8 = {shop_norm:.1f}")
                
                return shop_norm
            else:
                self.log_message(f"   [Ошибка] Нужно 2 даты, найдено {len(dates)}: {dates}")
                return None
                
        except Exception as e:
            self.log_message(f"   [Ошибка] При расчете нормы: {str(e)}")
            return None
        
    def detect_period(self):
        """Определяет отчетный период, рассчитывает норму магазина"""
        self.set_active_button("📅 Определить период")
        self.clear_log()
        self.log_message("📅 Определяю отчетный период из файлов...")
        
        # Сначала проверяем файлы
        if not self.manager.found_files:
            self.log_message("❌ Сначала проверьте файлы!")
            return
        
        # Определяем период
        if self.manager.detect_report_period():
            self.log_message(f"✅ Период определен: {self.manager.report_month}")
            self.log_message(f"   Полный период: {self.manager.report_period}")
            
            # 1. РАСЧЕТ НОРМЫ МАГАЗИНА
            shop_norm = self._calculate_shop_norm(self.manager.report_period)
            if shop_norm:
                self.shop_norm_hours = shop_norm
                self.log_message(f"⏱️ Норма часов 'Магазин': {self._format_russian_number(shop_norm, 1)}", color="green")
                self.log_message(f"   Формула: (дней в месяце / 7) × 5 × 8 = {shop_norm:.1f}")
            else:
                self.shop_norm_hours = None
                self.log_message("❌ Не удалось рассчитать норму для магазина")
                self.log_message("   Проверьте формат периода в файлах")
            
            # 2. СООБЩЕНИЕ О НОРМЕ ОФИСА
            self.log_message("\n" + "-"*70)
            self.log_message("⚙️  ТРЕБУЕТСЯ РУЧНАЯ НАСТРОЙКА:", color="red")
            self.log_message("")
            self.log_message("📋 Норма часов для отделов 'Офис' устанавливается вручную", color="red")
            self.log_message(f"⏱ Текущее значение: {self._format_russian_number(self.office_norm_hours, 0)}", color="green")
            self.log_message("")
            self.log_message("💡 Подсказка: норма часов офиса выставляется по производственному календарю")
            self.log_message("-"*60)
            
            # Обновляем метку периода
            self.period_label.config(text=f"Период: {self.manager.report_month}")
            
            # Подсвечиваем поле ввода нормы офиса
            self.office_norm_entry.config(background='#FFF3CD')  # Светло-желтый фон
            
            return True
        else:
            self.log_message("❌ Не удалось определить период!")
            self.log_message("   Проверьте файлы График.xls и Анализ продаж.xlsx")
            self.log_message("   В них должна быть строка формата: 'С 01.12.25 по 31.12.25'")
            return False
    
    def load_data(self):
        """Загружает данные"""
        self.set_active_button("📊 Загрузить данные")
        self.clear_log()
        self.log_message("📊 Загружаю данные из файлов. Этот процесс занимает немного времени.")
        
        # Проверяем период
        if not self.manager.report_period:
            self.log_message("⚠️  Сначала определите период!")
            return
        
        # Загружаем данные
        try:
            success = self.manager.load_all_data()
            if success:
                self.log_message("✅ Все данные успешно загружены.")
                
                # Показываем статистику
                self.log_message("\n📈 СТАТИСТИКА ЗАГРУЗКИ:")
                if self.manager.urs_data and self.manager.urs_data.get('success'):
                    filial_count = self.manager.urs_data.get('statistics', {}).get('unique_filials', 0)
                    self.log_message(f"  • Филиалов в файле УРС: {filial_count}")
                    dept_count = self.manager.urs_data.get('statistics', {}).get('departments_count', 0)
                    self.log_message(f"  • Отделов в файле УРС: {dept_count}")
                    
                    self.log_message(f"  --------")

                if self.manager.staff_data and self.manager.staff_data.get('success'):
                    summary = self.manager.staff_data['summary']
                    total_employees = summary.get('total_employees', 0)
                    total_branches = summary.get('total_branches', 0)
                    total_departments = summary.get('total_departments', 0)
                    
                    self.log_message(f"  • Филиалов в файле Сотрудники: {self._format_russian_number(total_branches)}")
                    self.log_message(f"  • Отделов в файле Сотрудники: {self._format_russian_number(total_departments)}")
                    self.log_message(f"  --------")
                    self.log_message(f"  • Продавцов в файле Сотрудники: {self._format_russian_number(total_employees)}")
                                
                if self.manager.schedule_data and 'error' not in self.manager.schedule_data:
                    self.log_message(f"  • Сотрудников в файле График: {len(self.manager.schedule_data['employees_df'])}")
                
                if self.manager.sales_data:
                    self.log_message(f"  • Продавцов в файле Анализ продаж: {len(self.manager.sales_data)}")
                    self.log_message(f"  --------")

                if self.manager.bonus_data and self.manager.bonus_data.get('success'):
                    self.log_message(f"  • Наименований бонусных товаров: {len(self.manager.bonus_data['bonus_items'])}")

                # Загрузка данных о заказных товарах
                zakaz_path = os.path.join("данные", "Заказ.xls")
                if os.path.exists(zakaz_path):
                    excluded = self.manager.urs_data.get('excluded_firms', []) if self.manager.urs_data.get('success') else []
                    zakaz_data = parse_zakaz_sales(zakaz_path, self.manager.staff_data, excluded)
                    self.manager.zakaz_data = zakaz_data
                    matched = zakaz_data.get('statistics', {}).get('matched_employees', 0)
                    self.log_message(f"  • Найдено сотрудников в Заказ.xls: {matched}", "green")
                else:
                    self.log_message("  ⚠️ Файл Заказ.xls не найден", "orange")
                    self.manager.zakaz_data = {'success': False, 'data': {}}
                
                if self.manager.bonus_data and self.manager.bonus_data.get('success'):
                    self.log_message(f"  • Наименований неликвидных товаров: {len(self.manager.bonus_data['non_liquid_items'])}")
                    self.log_message("")
                    self.log_message("-"*50)
                    self.log_message("")
                    self.log_message("⚙ Проверить совпадение кол-ва отделов и филиалов в файлах УРС и Сотрудники.", color="orange")
                    self.log_message("   Если данные разнятся, проверить данные в файлах.", color="orange")

                return True
            else:
                self.log_message("❌ Ошибка загрузки данных!")
                return False
                
        except Exception as e:
            self.log_message(f"❌ Ошибка при загрузке: {str(e)}")
            import traceback
            self.log_message(traceback.format_exc())
            return False
    
    def integrate_data(self):
        """Интегрирует данные"""
        self.set_active_button("🔄 Интегрировать данные")
        self.clear_log()
        self.log_message("🔄 Интегрирую данные из разных источников...")
        
        # Проверяем загружены ли данные
        if not self.manager.schedule_data or not self.manager.staff_data:
            self.log_message("❌ Сначала загрузите данные!", color="red")
            return
        
        try:
            from модули.data_integrator_simple import DataIntegrator
            
            # ПРОВЕРКА ДО интеграции
            if not hasattr(self, 'shop_norm_hours') or not self.shop_norm_hours:
                self.log_message("❌ НЕОБХОДИМО ПРЕДВАРИТЕЛЬНОЕ ДЕЙСТВИЕ:", color="red")
                self.log_message("="*60, color="orange")
                self.log_message("📋 Отделы 'Магазин' требуют расчета нормы часов", color="orange")
                self.log_message("", color="black")
                self.log_message("📌 ИНСТРУКЦИЯ:", color="blue")
                self.log_message("   1. Нажмите кнопку 'Определить период'", color="black")
                self.log_message("   2. Программа рассчитает норму для магазина", color="black")
                self.log_message("   3. Установите норму для офиса (если нужно)", color="black")
                self.log_message("   4. Нажмите 'Интегрировать данные' снова", color="black")
                self.log_message("", color="black")
                self.log_message("💡 Норма магазина рассчитывается автоматически", color="blue")
                self.log_message("   по формуле: (дней в месяце / 7) × 5 × 8", color="blue")
                self.log_message("="*60, color="orange")
                return False

            # Передаем shop_norm_hours в интегратор
            self.manager.shop_norm_hours = self.shop_norm_hours
            
            # Передаем office_norm_hours при интеграции
            self.manager.integrated_data = DataIntegrator.create_integrated_dataframe(
                self.manager,
                self.office_norm_hours  # норма офиса из интерфейса
            )
            
            if self.manager.integrated_data is not None:
                self.log_message("✅ Данные успешно интегрированы.")
                
                # Рассчитываем статистику интегрированных данных
                self.manager.calculate_integration_stats()
                
                # Показываем полную статистику
                stats = self.manager.integration_stats
                self.log_message("\n 📊 СТАТИСТИКА ИНТЕГРАЦИИ:")
                self.log_message(f"  • Всего записей: {self._format_russian_number(stats.get('total_records', 0))}")
                self.log_message(f"  • С графиком: {self._format_russian_number(stats.get('with_schedule', 0))}")
                self.log_message(f"  • Со структурой: {self._format_russian_number(stats.get('with_staff', 0))}")
                self.log_message(f"  • С продажами: {self._format_russian_number(stats.get('with_sales', 0))}")
                self.log_message(f"  • С обоими источниками: {self._format_russian_number(stats.get('with_both_sources', 0))}")
                self.log_message(f"  • Общая выручка: {self._format_russian_number(stats.get('total_revenue', 0))} руб.")
                self.log_message(f"  • Общая прибыль: {self._format_russian_number(stats.get('total_profit', 0))} руб.")
                self.log_message(f"  • Филиалов: {self._format_russian_number(stats.get('branches_count', 0))}")
                self.log_message(f"  • Отделов: {self._format_russian_number(stats.get('departments_count', 0))}")
                
                # Дополнительно: сотрудники без графика
                df = self.manager.integrated_data
                no_schedule = df[df['Источник_график'] == 'Нет']
                if not no_schedule.empty:
                    self.log_message(f"\n⚠️  Сотрудники без графика ({len(no_schedule)}):")
                    for _, row in no_schedule.iterrows():
                        self.log_message(f"   • {row['ФИО']} - {row['Отдел']}")
                
                # Показываем все отделы из интегрированных данных
                if stats.get('departments_list'):
                    dept_list = sorted(stats.get('departments_list', []))
                    self.log_message(f"\n📁 Все отделы в интегрированных данных ({len(dept_list)}):")
                    for dept in dept_list:
                        self.log_message(f"   • {dept}")
                
                # Проверка норм часов
                self.log_message(f"\n⏱️  Использованные нормы часов:")
                self.log_message(f"   • Магазин: {self._format_russian_number(self.shop_norm_hours, 1)} часов")
                self.log_message(f"   • Офис: {self._format_russian_number(self.office_norm_hours, 0)} часов")
                
                # Прокрутить к началу результатов
                self.info_text.see(1.0)
                
                return True
            else:
                self.log_message("❌ Ошибка интеграции данных!", color="red")
                return False
                
        except ValueError as e:
            # Обрабатываем нашу ошибку о неподсчитанной норме
            error_msg = str(e)
            if "норма часов" in error_msg.lower():
                self.log_message("❌ ОШИБКА ИНТЕГРАЦИИ:", color="red")
                self.log_message("="*60, color="orange")
                self.log_message("📋 Не рассчитана норма часов для магазина", color="orange")
                self.log_message("", color="black")
                self.log_message("📌 ВАШИ ДЕЙСТВИЯ:", color="blue")
                self.log_message("   1. Нажмите 'Определить период'", color="black")
                self.log_message("   2. Дождитесь расчета нормы магазина", color="black")
                self.log_message("   3. Нажмите 'Интегрировать данные' снова", color="black")
                self.log_message("="*60, color="orange")
            else:
                self.log_message(f"❌ Ошибка: {error_msg}", color="red")
            return False
                
        except Exception as e:
            self.log_message(f"❌ Ошибка при интеграции: {str(e)}", color="red")
            import traceback
            self.log_message(traceback.format_exc(), color="red")
            return False
    
    def calculate_salary(self):
        """Рассчитывает зарплату"""
        self.set_active_button("🧮 Рассчитать зарплату")
        self.clear_log()
        self.log_message("🧮 Рассчитываю зарплату по установленной логике...")
        
        # Проверяем интегрированы ли данные
        if self.manager.integrated_data is None or self.manager.integrated_data.empty:
            self.log_message("❌ Сначала интегрируйте данные!")
            return
        
        try:
            import importlib
            import модули.salary_calculator as salary_module
            importlib.reload(salary_module)
            from модули.salary_calculator import SalaryCalculator
            
            # Создаем калькулятор
            calculator = SalaryCalculator()
                      
            # ✅ ИСПРАВЛЕНО: передаем ТОЛЬКО 2 дополнительных аргумента (всего 3 с self)
            self.manager.calculations = calculator.calculate_salary(
                self.manager.integrated_data,
                self.office_norm_hours
            )
            
            if self.manager.calculations:
                self.log_message("✅ Расчет зарплаты выполнен успешно!")
                
                # Получаем DataFrame результатов
                results_df = self.manager.calculations['by_employee']
                
                # Информация о нормах часов
                shop_norm = results_df[results_df['Норма_часов'] != self.office_norm_hours]['Норма_часов'].unique()
                if len(shop_norm) > 0:
                    shop_norm_value = float(shop_norm[0])
                    if shop_norm_value.is_integer():
                        self.log_message(f"⏱️ Норма часов 'Магазин': {self._format_russian_number(shop_norm_value, 0)}")
                    else:
                        self.log_message(f"⏱️ Норма часов 'Магазин': {self._format_russian_number(shop_norm_value, 1)}")
                else:
                    self.log_message("⏱️ Норма часов 'Магазин': не рассчитана")
                self.log_message(f"⏱️ Норма часов 'Офис': {self._format_russian_number(self.office_norm_hours, 0)}")
                
                # Рассчитываем статистику из DataFrame
                total_salary = results_df['Зарплата_итого'].sum()
                avg_salary = results_df['Зарплата_итого'].mean()
                median_salary = results_df['Зарплата_итого'].median()
                total_hours = results_df['Отработано_часов'].sum()
                
                # Показываем краткую статистику
                self.log_message("\n📈 ИТОГИ РАСЧЕТА:")
                self.log_message(f"  • Сотрудников: {len(results_df)}")
                self.log_message(f"  • Отделов: {len(results_df['Отдел'].unique())}")
                self.log_message(f"  • Фонд зарплаты: {self._format_russian_number(total_salary)} руб.")
                self.log_message(f"  • Средняя зарплата: {self._format_russian_number(avg_salary)} руб.")
                self.log_message(f"  • Медианная зарплата: {self._format_russian_number(median_salary)} руб.")
                self.log_message(f"  • Всего часов: {self._format_russian_number(total_hours, 0)}")
                
                # Проверяем и выводим проблемы
                if ('problems' in self.manager.calculations and 
                    self.manager.calculations['problems'] is not None):
                    
                    problems = self.manager.calculations['problems']
                    
                    if problems['total_problems'] > 0:
                        self.log_message("\n" + "="*60, color="orange")
                        self.log_message("⚠️  ВНИМАНИЕ: ОБНАРУЖЕНЫ ПРОБЛЕМЫ С ДАННЫМИ", color="red")
                        self.log_message("="*60, color="orange")
                        self.log_message(f"Найдено проблем: {problems['total_problems']}", color="black")
                        self.log_message(f"Затронуто отделов: {problems['problem_departments']}", color="black")
                        self.log_message("", color="orange")
                        
                        # Показываем примеры проблем (первые 5)
                        self.log_message("Проблемы, требующие внимания:", color="red")
                        for i, problem in enumerate(problems['problem_list'][:5], 1):
                            self.log_message(f"{i}. {problem['ФИО']} - {problem['Отдел']}: {problem['Проблема']}", color="red")
                        
                        self.log_message("", color="orange")
                        self.log_message("💡 Рекомендация: Проверьте данные в файлах", color="black")
                        self.log_message("="*60, color="orange")

                # Тестируем сбор данных для отчетов
                if hasattr(self.manager, 'collect_all_indicators_for_reports'):
                    indicators = self.manager.collect_all_indicators_for_reports()
                    if indicators:
                        self.log_message(f"✅ Собрано показателей: {len(indicators)}", "green")
                
                return True
            else:
                self.log_message("❌ Ошибка расчета зарплаты!")
                return False
                
        except Exception as e:
            self.log_message(f"❌ Ошибка при расчете: {str(e)}")
            import traceback
            self.log_message(traceback.format_exc())
            return False

        
    
    def show_data_preview(self):
        """Показывает превью данных"""
        self.set_active_button("👁 Просмотреть данные")
        self.clear_log()
        self.log_message("👁 Просмотр данных...")
        
        # Создаем окно выбора
        dialog = tk.Toplevel(self.root)
        dialog.title("Просмотр данных")
        dialog.geometry("600x400")
        
        ttk.Label(dialog, text="Выберите какие данные просмотреть:", 
                 font=('Arial', 12)).pack(pady=10)
        
        options = [
            ("📅 График работы", self.preview_schedule),
            ("👥 Сотрудники", self.preview_staff),
            ("💰 Продажи", self.preview_sales),
            ("🎁 Бонусные товары", self.preview_bonus),
            ("⚙️ Настройки УРС", self.preview_urs),
            ("🔄 Интегрированные данные", self.preview_integrated),
            ("🧮 Результаты расчета", self.preview_calculations)
        ]
        
        for text, command in options:
            btn = ttk.Button(dialog, text=text, command=lambda c=command: self.run_preview(c, dialog),
                           width=30)
            btn.pack(pady=5)
        
        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)
    
    def run_preview(self, preview_func, dialog):
        """Запускает превью и закрывает диалог"""
        dialog.destroy()
        preview_func()

    def preview_schedule(self):
        """Превью данных графика - ВСЕХ сотрудников с часами"""
        self.clear_log()
        
        # Проверяем данные
        if not self.manager.schedule_data or 'error' in self.manager.schedule_data:
            self.log_message("❌ Данные графика не загружены!")
            return
        
        if not self.manager.staff_data or not self.manager.staff_data.get('success'):
            self.log_message("❌ Данные сотрудников не загружены!")
            return
        
        # Используем прямую вставку для заголовков таблицы
        self.info_text.insert(tk.END, "👥 ВСЕ СОТРУДНИКИ ИЗ ФАЙЛА 'СОТРУДНИКИ ПО ОТДЕЛАМ':\n", "black")
        self.info_text.insert(tk.END, f"Всего сотрудников: {self.manager.staff_data['summary']['total_employees']}\n", "black")
        self.info_text.insert(tk.END, f"Филиалов: {self.manager.staff_data['summary']['total_branches']}\n", "black")
        self.info_text.insert(tk.END, f"Отделов: {self.manager.staff_data['summary']['total_departments']}\n", "black")
        
        # 1. Создаем словарь с данными графика
        schedule_data = {}
        if self.manager.schedule_data and 'error' not in self.manager.schedule_data:
            schedule_df = self.manager.schedule_data['employees_df']
            for _, row in schedule_df.iterrows():
                fio_norm = self.normalize_name(row.get('ФИО', ''))
                schedule_data[fio_norm] = {
                    'Часы_всего': row.get('Часы_всего', 0),
                    'Выходные_дни': row.get('Выходные_дни', 0),  # Используем те же ключи, что в файле
                    'Отпуск_дни': row.get('Отпуск_дни', 0)
                }
        
        # 2. Собираем список всех сотрудников с часами
        employees_list = []
        
        for emp in self.manager.staff_data['employees']:
            fio = emp.get('ФИО', '')
            fio_norm = emp.get('ФИО_норм', self.normalize_name(fio))
            dept = emp.get('Отдел', 'Не указан')
            
            # Получаем данные из графика
            schedule_info = schedule_data.get(fio_norm, {})
            hours = schedule_info.get('Часы_всего', 0)
            weekend = schedule_info.get('Выходные_дни', 0)
            vacation = schedule_info.get('Отпуск_дни', 0)
            has_schedule = 'Да' if fio_norm in schedule_data else 'Нет'
            
            employees_list.append({
                '№': len(employees_list) + 1,
                'ФИО': fio,
                'Отдел': dept,
                'Часы': float(hours),
                'Вых': int(weekend),    # Ключ 'Вых'
                'Отп': int(vacation),   # Ключ 'Отп'
                'График': has_schedule  # Ключ 'График'
            })
        
        # Сортируем по ФИО
        employees_list.sort(key=lambda x: x['ФИО'])
        # Обновляем номера после сортировки
        for i, emp in enumerate(employees_list, 1):
            emp['№'] = i
        
        self.info_text.insert(tk.END, f"\n📋 ВСЕ СОТРУДНИКИ С ДАННЫМИ ИЗ ГРАФИКА ({len(employees_list)}):\n", "black")
        
        # 3. Формируем таблицу
        if employees_list:
            # Определяем ширины колонок (ключи должны совпадать с ключами в данных)
            widths = {
                '№': 4,
                'ФИО': 30,
                'Отдел': 32,
                'Часы': 5,
                'Вых': 3,      # Ключ 'Вых'
                'Отп': 3,      # Ключ 'Отп'
                'График': 7    # Ключ 'График'
            }
            
            # Создаем заголовок БЕЗ временной метки
            header_parts = []
            for key, width in widths.items():
                # Красивые названия для заголовков
                display_key = {
                    'Вых': 'Вых',
                    'Отп': 'Отп',
                    'График': 'График'
                }.get(key, key)
                header_parts.append(f"{display_key:{width}}")
            
            header_line = " ".join(header_parts)
            self.info_text.insert(tk.END, header_line + "\n", "black")
            
            # Разделительная линия
            separator_line = "-" * len(header_line)
            self.info_text.insert(tk.END, separator_line + "\n", "black")
            
            # Данные
            for emp in employees_list:
                row_parts = []
                
                # Номер
                row_parts.append(f"{emp['№']:{widths['№']}}")
                
                # ФИО
                fio = emp['ФИО']
                if len(fio) > widths['ФИО'] - 2:
                    fio = fio[:widths['ФИО']-3] + "..."
                row_parts.append(f"{fio:{widths['ФИО']}}")
                
                # Отдел
                dept = emp['Отдел']
                if len(dept) > widths['Отдел'] - 2:
                    dept = dept[:widths['Отдел']-3] + "..."
                row_parts.append(f"{dept:{widths['Отдел']}}")
                
                # Часы
                hours = emp['Часы']
                hours_str = f"{hours:.1f}" if hours != int(hours) else str(int(hours))
                row_parts.append(f"{hours_str:>{widths['Часы']}}")
                
                # Выходные (используем ключ 'Вых')
                row_parts.append(f"{emp['Вых']:>{widths['Вых']}}")
                
                # Отпуск (используем ключ 'Отп')
                row_parts.append(f"{emp['Отп']:>{widths['Отп']}}")
                
                # Статус графика (используем ключ 'График')
                status = emp['График']
                color = "green" if status == 'Да' else "orange"
                row_parts.append(f"{status:^{widths['График']}}")
                
                row_line = " ".join(row_parts)
                self.info_text.insert(tk.END, row_line + "\n", color)
            
            # Статистика
            self.info_text.insert(tk.END, f"\n📊 СТАТИСТИКА:\n", "black")
            
            has_schedule = sum(1 for emp in employees_list if emp['График'] == 'Да')
            no_schedule = sum(1 for emp in employees_list if emp['График'] == 'Нет')
            total_hours = sum(emp['Часы'] for emp in employees_list)
            total_weekend = sum(emp['Вых'] for emp in employees_list)
            total_vacation = sum(emp['Отп'] for emp in employees_list)
            
            self.info_text.insert(tk.END, f"• С графиком: {has_schedule} сотрудников\n", "black")
            self.info_text.insert(tk.END, f"• Без графика: {no_schedule} сотрудников\n", "black")
            self.info_text.insert(tk.END, f"• Всего часов: {total_hours:.1f}\n", "black")
            self.info_text.insert(tk.END, f"• Всего выходных дней: {total_weekend}\n", "black")
            self.info_text.insert(tk.END, f"• Всего отпускных дней: {total_vacation}\n", "black")
            
            # Сотрудники без графика
            if no_schedule > 0:
                self.info_text.insert(tk.END, f"\n⚠️ СОТРУДНИКИ БЕЗ ГРАФИКА:\n", "orange")
                no_schedule_list = [emp for emp in employees_list if emp['График'] == 'Нет']
                for i, emp in enumerate(no_schedule_list[:20], 1):
                    self.info_text.insert(tk.END, f"  {i:2}. {emp['ФИО'][:30]} - {emp['Отдел'][:20]}\n", "orange")
                if len(no_schedule_list) > 20:
                    self.info_text.insert(tk.END, f"  ... и еще {len(no_schedule_list) - 20} сотрудников\n", "orange")
            
            self.info_text.see(1.0)

    def normalize_name(self, full_name):
        """Нормализует ФИО для сравнения"""
        if not full_name or not isinstance(full_name, str):
            return ''
        return ' '.join(full_name.strip().split()).upper()
    
    def preview_staff(self):
        """Превью данных сотрудников"""
        self.clear_log()
        if self.manager.staff_data and self.manager.staff_data.get('success'):
            self.log_message("👥 ДАННЫЕ СОТРУДНИКОВ:")
            self.log_message(f"Всего сотрудников: {self.manager.staff_data['summary']['total_employees']}")
            self.log_message(f"Филиалов: {self.manager.staff_data['summary']['total_branches']}")
            self.log_message(f"Отделов: {self.manager.staff_data['summary']['total_departments']}")
            
            # Примеры сотрудников
            employees = self.manager.staff_data['employees'][:10]
            self.log_message("\nПримеры сотрудников (первые 10):")
            for i, emp in enumerate(employees, 1):
                self.log_message(f"{i:2}. {emp['ФИО'][:30]:30} | {emp['Филиал'][:15]:15} | {emp['Отдел'][:20]:20}")
        else:
            self.log_message("❌ Данные сотрудников не загружены!")
    
    def preview_sales(self):
        """Превью данных продаж"""
        self.clear_log()
        if self.manager.sales_data:
            self.log_message("💰 ДАННЫЕ ПРОДАЖ:")
            self.log_message(f"Продавцов: {len(self.manager.sales_data)}")
            
            # Первые 5 продавцов
            count = 0
            for seller, data in self.manager.sales_data.items():
                if count >= 5:
                    break
                self.log_message(f"\n• {seller}:")
                self.log_message(f"  Выручка: {self._format_russian_number(data.get('total_revenue', 0))} руб.")
                self.log_message(f"  Прибыль: {self._format_russian_number(data.get('total_profit', 0))} руб.")
                count += 1
        else:
            self.log_message("❌ Данные продаж не загружены!")
    
    def preview_bonus(self):
        """Превью бонусных товаров"""
        self.clear_log()
        if self.manager.bonus_data and self.manager.bonus_data.get('success'):
            stats = self.manager.bonus_data['statistics']
            self.log_message("🎁 БОНУСНЫЕ ТОВАРЫ:")
            self.log_message(f"Бонусных товаров: {stats['bonus_count']}")
            self.log_message(f"Неликвидов: {stats['non_liquid_count']}")
            self.log_message(f"Всего товаров: {stats['total_unique']}")
            
            # Примеры бонусных товаров
            if self.manager.bonus_data['bonus_items']:
                self.log_message("\nПримеры бонусных товаров (первые 10):")
                for i, code in enumerate(list(self.manager.bonus_data['bonus_items'])[:10], 1):
                    info = self.manager.bonus_data['items_info'].get(code, {})
                    name = info.get('название', 'Нет названия')[:30]
                    self.log_message(f"{i:2}. {code} - {name}...")
        else:
            self.log_message("❌ Данные бонусов не загружены!")
    
    def preview_urs(self):
        """Превью настроек УРС"""
        self.clear_log()
        if self.manager.urs_data and self.manager.urs_data.get('success'):
            stats = self.manager.urs_data['statistics']
            self.log_message("⚙️ НАСТРОЙКИ УРС:")
            self.log_message(f"Отделов: {stats['unique_departments']}")
            self.log_message(f"Филиалов: {stats['unique_filials']}")
            
            # Примеры настроек
            self.log_message("\nПримеры отделов (первые 10):")
            for i, (dept, settings) in enumerate(list(self.manager.urs_data['departments'].items())[:10], 1):
                oklad = settings.get('оклад', 0)
                minim = settings.get('минималка', 0)
                self.log_message(f"{i:2}. {dept[:30]:30} | Оклад: {self._format_russian_number(oklad, 0):8} | Мин.: {self._format_russian_number(minim, 0):8}")
        else:
            self.log_message("❌ Данные УРС не загружены!")
    
    def preview_integrated(self):
        """Превью интегрированных данные"""
        self.clear_log()
        if self.manager.integrated_data is not None:
            df = self.manager.integrated_data
            self.log_message("🔄 ИНТЕГРИРОВАННЫЕ ДАННЫЕ:")
            self.log_message(f"Всего записей: {len(df)}")
            self.log_message(f"С графиком: {(df['Источник_график'] == 'Да').sum()}")
            self.log_message(f"Со структурой: {(df['Источник_сотрудники'] == 'Да').sum()}")
            self.log_message(f"С продажами: {(df['Выручка'] > 0).sum()}")
            self.log_message(f"Общая выручка: {self._format_russian_number(df['Выручка'].sum())} руб.")
            
            # Показываем первые 10 записей
            self.log_message("\nПервые 10 записей:")
            preview_cols = ['ФИО', 'Филиал', 'Отдел', 'Часы_всего', 'Выручка']
            preview_cols = [col for col in preview_cols if col in df.columns]
            
            preview_df = df[preview_cols].head(10)
            self.log_message(preview_df.to_string(index=False))
        else:
            self.log_message("❌ Интегрированные данные не созданы!")
    
    def preview_calculations(self):
        """Превью результатов расчета"""
        self.clear_log()
        if not hasattr(self.manager, 'calculations') or not self.manager.calculations:
            self.log_message("❌ Расчет не выполнен!")
            return

        try:
            # Проверяем наличие проблем
            if ('problems' in self.manager.calculations and 
                self.manager.calculations['problems'] is not None):
                
                problems = self.manager.calculations['problems']
                
                self.log_message("⚠️ ⚠️ ⚠️  ВНИМАНИЕ: НАЙДЕНЫ ПРОБЛЕМЫ С ДАННЫМИ ⚠️ ⚠️ ⚠️", color="red")
                self.log_message("="*80, color="red")
                self.log_message(f"Всего проблем: {problems['total_problems']}", color="red")
                self.log_message(f"Отделов с проблемами: {problems.get('problem_departments', 0)}", color="red")
                
                # Показываем примеры проблем (первые 10)
                if problems['problem_list']:
                    self.log_message("\nПримеры проблем:", color="orange")
                    for problem in problems['problem_list'][:10]:
                        self.log_message(f"• {problem['ФИО']} - {problem['Отдел']}: {problem['Проблема']}", color="orange")
                
                self.log_message("="*80, color="red")
                self.log_message("💡 РЕКОМЕНДАЦИЯ: Проверьте данные в файлах перед утверждением расчета", color="red")
                self.log_message("", color="black")
            
        except Exception as e:
            self.log_message(f"⚠️ Ошибка при проверке проблем: {str(e)}", color="orange")
            
        try:
            # Получаем данные расчета
            results_df = self.manager.calculations['by_employee']
            
            # Выводим отчет в текстовое поле
            self.log_message("🧮 РЕЗУЛЬТАТЫ РАСЧЕТА ЗАРПЛАТЫ")
            self.log_message("="*60)
            
            # Создаем summary если его нет
            if 'summary' in self.manager.calculations:
                summary = self.manager.calculations['summary']
            else:
                # Создаем summary из данных
                summary = {
                    'total_employees': len(results_df),
                    'total_departments': results_df['Отдел'].nunique(),
                    'total_salary': results_df['Зарплата_итого'].sum(),
                    'avg_salary': results_df['Зарплата_итого'].mean(),
                    'total_hours': results_df['Отработано_часов'].sum(),
                    'total_sales': results_df['Выручка_всего'].sum() if 'Выручка_всего' in results_df.columns else 0
                }
            
            self.log_message(f"\n📊 ОБЩАЯ СТАТИСТИКА:")
            self.log_message(f"  Сотрудников: {summary['total_employees']}")
            self.log_message(f"  Отделов: {summary['total_departments']}")
            self.log_message(f"  Фонд зарплаты: {self._format_russian_number(summary['total_salary'])} руб.")
            self.log_message(f"  Средняя зарплата: {self._format_russian_number(summary['avg_salary'])} руб.")
            
            # Сортировка по отделам и рейтингу (убывание)
            results_df = results_df.sort_values(['Отдел', 'Рейтинг'], ascending=[True, False])
            
            self.log_message("\n" + "="*80)
            self.log_message("РАСПРЕДЕЛЕНИЕ ЗАРПЛАТЫ ПО ОТДЕЛАМ (Топ по рейтингу)")
            self.log_message("="*80)
            
            # Группировка по отделам
            current_dept = None
            emp_num = 0
            
            for _, row in results_df.iterrows():
                if row['Отдел'] != current_dept:
                    current_dept = row['Отдел']
                    emp_num = 0
                    self.log_message(f"\n🏢 ОТДЕЛ: {current_dept}")
                    self.log_message("-" * 60)
                    
                    # Заголовок таблицы для отдела
                    self.log_message(f"{'№':3} {'ФИО':30} {'Зарплата':>12} {'Компоненты':40}")
                    self.log_message("-" * 60)
                
                emp_num += 1
                total_salary = row['Зарплата_итого']
                
                # Получаем данные из НОВОЙ структуры
                dolya_kotla = row.get('Доля_котла', 0)
                okladnaya_chast = row.get('Окладная_часть', 0)
                minimalka_ind = row.get('Минималка_инд', 0)
                primenena_garantiya = row.get('Применена_гарантия', 0)
                mesto = row.get('Место', 0)
                
                # Форматируем имя
                fio_parts = row['ФИО'].split()
                if len(fio_parts) >= 2:
                    short_name = f"{fio_parts[0]} {fio_parts[1][0]}.{fio_parts[2][0] if len(fio_parts) > 2 else ''}"
                else:
                    short_name = row['ФИО']
                
                # Добавляем значок места если есть гарантия
                mesto_symbol = ""
                if mesto > 0:
                    mesto_symbol = f" [{mesto}🏆]"
                
                # Проверяем, получил ли сотрудник минималку
                if abs(total_salary - minimalka_ind) < 1:  # Если зарплата равна минималке
                    actual_components = f"Минималка: {self._format_russian_number(minimalka_ind)}"
                else:
                    actual_components = f"Котел: {self._format_russian_number(dolya_kotla)} + Оклад: {self._format_russian_number(okladnaya_chast)}"
                    
                    # Добавляем гарантию, если она была применена
                    if primenena_garantiya > 0:
                        actual_components += f" + Гарантия: {self._format_russian_number(primenena_garantiya)}"
                
                # Формируем строку результата
                zarplata_formatted = self._format_russian_number(total_salary)
                result_line = f"{emp_num:3}. {short_name[:28]:30} {zarplata_formatted:>12} руб. {actual_components}{mesto_symbol}"
                
                # Выделяем цветом
                color = "green" if mesto > 0 else "black"
                self.log_message(result_line, color=color)
            
            # Итоги по отделам
            self.log_message("\n" + "="*80)
            self.log_message("ИТОГИ ПО ОТДЕЛАМ")
            self.log_message("="*80)
            
            dept_summary = {}
            for dept, group in results_df.groupby('Отдел'):
                dept_summary[dept] = {
                    'сотрудников': len(group),
                    'общая_зарплата': group['Зарплата_итого'].sum(),
                    'котёл': group['Доля_котла'].sum(),
                    'оклад': group['Окладная_часть'].sum(),
                    'гарантии': group['Применена_гарантия'].sum(),
                    'минималка': group[abs(group['Зарплата_итого'] - group['Минималка_инд']) < 1]['Зарплата_итого'].sum(),
                    'продажи': group['Выручка_всего'].sum() if 'Выручка_всего' in group.columns else 0
                }
            
            # Сортировка отделов по общей зарплате
            for dept in sorted(dept_summary.keys(), key=lambda x: dept_summary[x]['общая_зарплата'], reverse=True):
                data = dept_summary[dept]
                self.log_message(f"\n{dept}:")
                self.log_message(f"  Сотрудников: {data['сотрудников']}")
                self.log_message(f"  Общая зарплата: {self._format_russian_number(data['общая_зарплата'])} руб.")
                self.log_message(f"    • Из котла: {self._format_russian_number(data['котёл'])} руб.")
                self.log_message(f"    • Окладная часть: {self._format_russian_number(data['оклад'])} руб.")
                if data['гарантии'] > 0:
                    self.log_message(f"    • Гарантии лидеров: {self._format_russian_number(data['гарантии'])} руб.")
                if data['минималка'] > 0:
                    self.log_message(f"    • Минималка: {self._format_russian_number(data['минималка'])} руб.")
                self.log_message(f"  Выручка отдела: {self._format_russian_number(data['продажи'])} руб.")
            
            # Показываем топ-3 по рейтингу
            top3 = results_df.nlargest(3, 'Рейтинг')
            if not top3.empty:
                self.log_message("\n" + "="*80)
                self.log_message("🏆 ТОП-3 ПО РЕЙТИНГУ")
                self.log_message("="*80)
                
                for i, (_, row) in enumerate(top3.iterrows(), 1):
                    rating_percent = row['Рейтинг'] * 100
                    self.log_message(f"{i}. {row['ФИО']} - {row['Отдел']}")
                    self.log_message(f"   Рейтинг: {rating_percent:.1f}% | "
                                   f"Зарплата: {self._format_russian_number(row['Зарплата_итого'])} руб.")
            
            # Показываем примененные гарантии
            with_guarantees = results_df[results_df['Применена_гарантия'] > 0]
            if not with_guarantees.empty:
                self.log_message("\n" + "="*80)
                self.log_message("🎯 СОТРУДНИКИ С ГАРАНТИЯМИ")
                self.log_message("="*80)
                
                for _, row in with_guarantees.iterrows():
                    self.log_message(f"• {row['ФИО']} - {row['Отдел']} ({row['Место']} место)")
                    self.log_message(f"  Было: {self._format_russian_number(row['Зарплата_предв'])} руб. → "
                                   f"Стало: {self._format_russian_number(row['Зарплата_итого'])} руб. "
                                   f"(+{self._format_russian_number(row['Применена_гарантия'] - row['Зарплата_предв'])} руб.)")
            
            # Дополнительная статистика
            self.log_message("\n" + "="*80)
            self.log_message("📈 ДОПОЛНИТЕЛЬНАЯ СТАТИСТИКА")
            self.log_message("="*80)
            
            avg_hours = results_df['Отработано_часов'].mean()
            avg_norm = results_df['Норма_часов'].mean()
            percent_norm = (avg_hours / avg_norm * 100) if avg_norm > 0 else 0
            
            self.log_message(f"Средняя выработка: {avg_hours:.1f} часов ({percent_norm:.1f}% от нормы)")
            self.log_message(f"Минимальная зарплата: {self._format_russian_number(results_df['Зарплата_итого'].min())} руб.")
            self.log_message(f"Максимальная зарплата: {self._format_russian_number(results_df['Зарплата_итого'].max())} руб.")
            
        except Exception as e:
            self.log_message(f"❌ Ошибка при показе результатов: {str(e)}")
            import traceback
            self.log_message(traceback.format_exc())
   
    
    def show_dashboard(self):
        """Генерация профессионального HTML дашборда (точный дизайн разработчика)"""
        self.set_active_button("📈 Дашборд")
        
        if not hasattr(self.manager, 'calculations') or not self.manager.calculations:
            self.clear_log()
            self.log_message("❌ Сначала выполните расчет зарплаты!", "red")
            return
        
        try:
            self.clear_log()
            self.log_message("📊 Генерация профессионального дашборда...", "blue")
            
            # Используем новый профессиональный дашборд
            from модули.manager_dashboard_pro import ManagerDashboardPro
            
            # Создаем дашборд
            generator = ManagerDashboardPro(self.manager)
            filepath = generator.generate()
            
            self.log_message(f"✅ Профессиональный дашборд создан!", "green")
            self.log_message(f"📂 Файл: {os.path.basename(filepath)}", "blue")
            self.log_message(f"📁 Папка: {os.path.dirname(filepath)}", "blue")
            
            # Показать сообщение с опциями
            import webbrowser
            result = messagebox.askyesno(
                "Дашборд создан", 
                f"Профессиональный дашборд успешно создан!\n\n"
                f"Файл: {os.path.basename(filepath)}\n"
                f"Путь: {filepath}\n\n"
                f"Открыть в браузере?",
                parent=self.root
            )
            
            if result:
                # Открыть в браузере по умолчанию
                webbrowser.open(f"file://{os.path.abspath(filepath)}")
                self.log_message("🌐 Дашборд открыт в браузере", "green")
            
            # Показать пароли для доступа
            self.log_message("\n🔐 ПАРОЛИ ДЛЯ ДОСТУПА:", "purple")
            self.log_message("• Мастер-пароль (директор): MASTER_KEY", "black")
            self.log_message("• Филиал БД1: BD1_PASS", "black")
            self.log_message("• Филиал БД3: BD3_PASS", "black")
            self.log_message("• Филиал БД4: BD4_PASS", "black")
            
            # Предупреждение о данных
            self.log_message("\n⚠️  ВНИМАНИЕ:", "orange")
            self.log_message("• Данные о выходных/отпуске/больничных берутся из графика", "orange")
            self.log_message("• Убедитесь, что файл График.xls содержит эти данные", "orange")
                
        except Exception as e:
            self.log_message(f"❌ Ошибка создания дашборда: {str(e)}", "red")
            import traceback
            self.log_message(traceback.format_exc(), "orange")
        
    
    
    def save_results(self):
        """Сохраняет результаты"""
        self.set_active_button("💾 Сохранить результаты")
        self.clear_log()
        self.log_message("💾 Сохранение результатов...")

    def create_report(self):
        """Создает отчет в Excel"""
        self.set_active_button("📄 Создать отчет Excel")
        self.clear_log()
        
        if not hasattr(self.manager, 'calculations') or not self.manager.calculations:
            self.log_message("❌ Сначала рассчитайте зарплату!", "red")
            return
        
        try:
            # ИСПРАВЬ ЭТО: используем SimpleReportGenerator
            from модули.simple_report import SimpleReportGenerator
            
            self.log_message("📋 Создание отчета...", "blue")
            
            generator = SimpleReportGenerator(self.manager)
            filepath, message = generator.create_salary_report()
            
            if filepath:
                self.log_message(f"✅ {message}", "green")
                self.log_message(f"📂 Файл сохранен: {filepath}", "blue")
            else:
                self.log_message(f"❌ {message}", "red")
                
        except Exception as e:
            self.log_message(f"❌ Ошибка: {str(e)}", "red")

    def create_simple_report(self):
        """Создает простой отчет по расчетам"""
        self.set_active_button("📋 Простой отчет")
        self.clear_log()
        
        if not hasattr(self.manager, 'calculations') or not self.manager.calculations:
            self.log_message("❌ Сначала рассчитайте зарплату!", "red")
            return
        
        try:
            from модули.simple_report import SimpleReportGenerator
            
            self.log_message("📋 Создание простого отчета...", "blue")
            
            generator = SimpleReportGenerator(self.manager)
            filepath, message = generator.create_salary_report()
            
            if filepath:
                self.log_message(f"✅ {message}", "green")
                self.log_message(f"📂 Файл сохранен: {filepath}", "blue")
            else:
                self.log_message(f"❌ {message}", "red")
                
        except Exception as e:
            self.log_message(f"❌ Ошибка: {str(e)}", "red")
    
    def update_office_norm(self):
        """Обновляет норму часов для офисных отделов"""
        try:
            self.office_norm_hours = int(self.office_norm_entry.get())
            self.log_message(f"✅ Норма часов офиса установлена: {self._format_russian_number(self.office_norm_hours, 0)}")
            
            # Если есть данные УРС, применяем норму
            if self.manager.urs_data and self.manager.urs_data.get('success'):
                from модули.parse_urs_integrated import apply_office_norm_hours
                self.manager.urs_data = apply_office_norm_hours(
                    self.manager.urs_data, 
                    self.office_norm_hours
                )
                self.log_message("✅ Норма применена к отделам 'Офис'")
                
        except ValueError:
            self.log_message("❌ Ошибка: введите целое число")
            self.office_norm_entry.delete(0, tk.END)
            self.office_norm_entry.insert(0, str(self.office_norm_hours))

def main():
    root = tk.Tk()
    app = SalaryCalculatorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
