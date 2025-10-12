# main.py
import os
import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date
import threading
import queue
from pathlib import Path
import re
import sys
import importlib
import inspect
import shutil
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PIL import Image, ImageTk
import sqlite3

# Версия программы
VERSION = "3.9.5"

# Проверяем наличие tksheet
try:
    import tksheet
    TKSHEET_AVAILABLE = True
except ImportError:
    TKSHEET_AVAILABLE = False
    logging.error("Библиотека tksheet не установлена. Отчет будет ограничен в функциях.")

class QueueHandler(logging.Handler):
    """Кастомный обработчик логов для отправки сообщений в очередь"""
    
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
    
    def emit(self, record):
        self.log_queue.put(self.format(record))

class DatabaseManager:
    """Менеджер базы данных для хранения истории переименований"""
    
    def __init__(self, db_file="rename_history.db"):
        self.db_file = db_file
        self.init_database()
    
    def init_database(self):
        """Инициализация базы данных"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # Создаем таблицу для истории переименований
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS rename_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    create_date TEXT NOT NULL,
                    route TEXT NOT NULL,
                    original_name TEXT NOT NULL,
                    new_name TEXT NOT NULL,
                    file_path TEXT NOT NULL
                )
            ''')
            
            # Создаем индекс для быстрого поиска по дате
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_date 
                ON rename_history(create_date)
            ''')
            
            # Создаем индекс для быстрого поиска по маршруту
            cursor.execute('''
                CREATE INDEX IF NOT EXISTS idx_route 
                ON rename_history(route)
            ''')
            
            conn.commit()
            conn.close()
            logging.info("База данных инициализирована успешно")
            
        except Exception as e:
            logging.error(f"Ошибка инициализации базы данных: {e}")
    
    def add_record(self, timestamp, route, original_name, new_name, file_path):
        """Добавление записи в базу данных"""
        try:
            create_date = datetime.now().strftime("%Y-%m-%d")
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO rename_history 
                (timestamp, create_date, route, original_name, new_name, file_path)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (timestamp, create_date, route, original_name, new_name, file_path))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logging.error(f"Ошибка добавления записи в базу данных: {e}")
            return False
    
    def get_records_by_date(self, target_date=None):
        """Получение записей за определенную дату"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if target_date:
                cursor.execute('''
                    SELECT * FROM rename_history 
                    WHERE create_date = ? 
                    ORDER BY timestamp DESC
                ''', (target_date,))
            else:
                cursor.execute('''
                    SELECT * FROM rename_history 
                    ORDER BY timestamp DESC
                ''')
            
            records = cursor.fetchall()
            conn.close()
            
            # Преобразуем в список словарей
            columns = ['id', 'timestamp', 'create_date', 'route', 'original_name', 'new_name', 'file_path']
            result = []
            for record in records:
                result.append(dict(zip(columns, record)))
            
            return result
        except Exception as e:
            logging.error(f"Ошибка получения записей из базы данных: {e}")
            return []
    
    def get_all_dates(self):
        """Получение всех уникальных дат из базы данных"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT DISTINCT create_date FROM rename_history 
                ORDER BY create_date DESC
            ''')
            
            dates = [row[0] for row in cursor.fetchall()]
            conn.close()
            return dates
        except Exception as e:
            logging.error(f"Ошибка получения дат из базы данных: {e}")
            return []
    
    def clear_records_by_date(self, target_date):
        """Удаление записей за определенную дату"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                DELETE FROM rename_history 
                WHERE create_date = ?
            ''', (target_date,))
            
            conn.commit()
            conn.close()
            return cursor.rowcount
        except Exception as e:
            logging.error(f"Ошибка удаления записей из базы данных: {e}")
            return 0
    
    def clear_all_records(self):
        """Удаление всех записей"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM rename_history')
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logging.error(f"Ошибка очистки базы данных: {e}")
            return False

class TemplateBuilderDialog:
    """Диалоговое окно для визуального построения шаблона"""
    
    def __init__(self, parent, current_template, settings):
        self.parent = parent
        self.current_template = current_template
        self.settings = settings
        self.result_template = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Конструктор шаблона имени файла")
        self.dialog.geometry("700x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        # Центрирование диалога
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.create_widgets()
        self.update_preview()
    
    def create_widgets(self):
        """Создание виджетов диалога"""
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Левая часть - доступные переменные
        left_frame = ttk.LabelFrame(main_frame, text="Доступные переменные", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Правая часть - предпросмотр и управление
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Список переменных
        variables_frame = ttk.Frame(left_frame)
        variables_frame.pack(fill=tk.BOTH, expand=True)
        
        # Группы переменных
        self.create_variable_group(variables_frame, "Основные переменные:", [
            ("{project}", "Проект", self.settings.settings["project"]),
            ("{CN}", "Тип ЦН", self.settings.settings["cn_type"]),
            ("{route}", "Маршрут", self.settings.settings["route"]),
            ("{date}", "Дата", self.format_date_example()),
            ("{counter}", "Счетчик", "001"),
            ("{extension}", "Расширение файла", "jpg")
        ])
        
        self.create_variable_group(variables_frame, "Пользовательские переменные:", [
            ("{1}", "Переменная 1", self.settings.settings["var1"]),
            ("{2}", "Переменная 2", self.settings.settings["var2"]),
            ("{3}", "Переменная 3", self.settings.settings["var3"])
        ])
        
        # Разделитель
        ttk.Separator(variables_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # Статические элементы
        static_frame = ttk.Frame(variables_frame)
        static_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(static_frame, text="Статические элементы:").pack(anchor=tk.W)
        
        static_buttons_frame = ttk.Frame(static_frame)
        static_buttons_frame.pack(fill=tk.X, pady=5)
        
        for text in ["-", "_", "(", ")", "[", "]"]:
            btn = ttk.Button(static_buttons_frame, text=text, width=3,
                           command=lambda t=text: self.insert_text(t))
            btn.pack(side=tk.LEFT, padx=2)
        
        # Правая часть - редактирование шаблона
        edit_frame = ttk.LabelFrame(right_frame, text="Редактирование шаблона", padding="10")
        edit_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Текущий шаблон
        ttk.Label(edit_frame, text="Текущий шаблон:").pack(anchor=tk.W)
        
        self.template_text = tk.Text(edit_frame, height=3, wrap=tk.WORD)
        self.template_text.pack(fill=tk.X, pady=5)
        self.template_text.insert("1.0", self.current_template)
        
        # Привязываем событие изменения текста
        self.template_text.bind("<KeyRelease>", self.on_template_change)
        
        # Кнопки для быстрого редактирования
        quick_buttons_frame = ttk.Frame(edit_frame)
        quick_buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(quick_buttons_frame, text="Очистить", 
                  command=self.clear_template).pack(side=tk.LEFT, padx=2)
        ttk.Button(quick_buttons_frame, text="Сбросить к стандарту", 
                  command=self.reset_to_default).pack(side=tk.LEFT, padx=2)
        
        # Предпросмотр
        preview_frame = ttk.LabelFrame(right_frame, text="Предпросмотр имени файла", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        self.preview_var = tk.StringVar(value="Пример: project_20231201_route_001_CN.jpg")
        preview_label = ttk.Label(preview_frame, textvariable=self.preview_var, 
                                 wraplength=400, justify=tk.LEFT)
        preview_label.pack(fill=tk.BOTH, expand=True)
        
        # История шаблонов
        history_frame = ttk.LabelFrame(right_frame, text="История шаблонов", padding="10")
        history_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.history_var = tk.StringVar()
        history_combo = ttk.Combobox(history_frame, textvariable=self.history_var,
                                    values=self.settings.settings.get("template_history", []))
        history_combo.pack(fill=tk.X, pady=5)
        history_combo.bind("<<ComboboxSelected>>", self.on_history_select)
        
        # Кнопки управления
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Сохранить", 
                  command=self.save_template).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Применить", 
                  command=self.apply_template).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Отмена", 
                  command=self.dialog.destroy).pack(side=tk.RIGHT, padx=5)
    
    def format_date_example(self):
        """Возвращает пример даты в текущем формате"""
        date_format = self.settings.settings.get("date_format", "ГГГГММДД")
        return self.format_date_by_format(datetime.now(), date_format)
    
    def format_date_by_format(self, date_obj, date_format):
        """Форматирует дату по выбранному формату"""
        format_mapping = {
            "ДДММГГГГ": date_obj.strftime("%d%m%Y"),
            "ДДММГГ": date_obj.strftime("%d%m%y"),
            "ГГГГММДД": date_obj.strftime("%Y%m%d"),
            "ДД.ММ.ГГГГ": date_obj.strftime("%d.%m.%Y"),
            "ДД.ММ.ГГ": date_obj.strftime("%d.%m.%y"),
            "ГГГГ.ММ.ДД": date_obj.strftime("%Y.%m.%d")
        }
        return format_mapping.get(date_format, date_obj.strftime("%Y%m%d"))
    
    def create_variable_group(self, parent, title, variables):
        """Создание группы переменных"""
        group_frame = ttk.Frame(parent)
        group_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(group_frame, text=title, font=('Arial', 9, 'bold')).pack(anchor=tk.W)
        
        for var_code, var_name, example_value in variables:
            var_frame = ttk.Frame(group_frame)
            var_frame.pack(fill=tk.X, pady=2)
            
            # Кнопка для вставки переменной
            btn = ttk.Button(var_frame, text=var_code, width=10,
                           command=lambda v=var_code: self.insert_variable(v))
            btn.pack(side=tk.LEFT, padx=(0, 5))
            
            # Описание переменной
            desc_text = f"{var_name} (пример: {example_value})"
            ttk.Label(var_frame, text=desc_text, font=('Arial', 8)).pack(side=tk.LEFT)
    
    def insert_variable(self, variable):
        """Вставка переменной в шаблон"""
        self.template_text.insert(tk.INSERT, variable)
        self.update_preview()
    
    def insert_text(self, text):
        """Вставка статического текста"""
        self.template_text.insert(tk.INSERT, text)
        self.update_preview()
    
    def clear_template(self):
        """Очистка шаблона"""
        self.template_text.delete("1.0", tk.END)
        self.update_preview()
    
    def reset_to_default(self):
        """Сброс к стандартному шаблону"""
        default_template = "{project}_{date}_{route}_{counter}_{CN}"
        self.template_text.delete("1.0", tk.END)
        self.template_text.insert("1.0", default_template)
        self.update_preview()
    
    def on_template_change(self, event=None):
        """Обработка изменения шаблона"""
        self.update_preview()
    
    def on_history_select(self, event=None):
        """Выбор шаблона из истории"""
        selected_template = self.history_var.get()
        if selected_template:
            self.template_text.delete("1.0", tk.END)
            self.template_text.insert("1.0", selected_template)
            self.update_preview()
    
    def update_preview(self):
        """Обновление предпросмотра"""
        try:
            template = self.template_text.get("1.0", tk.END).strip()
            
            # Заменяем переменные на примеры значений
            preview = template
            preview = preview.replace("{project}", self.settings.settings["project"])
            preview = preview.replace("{CN}", self.settings.settings["cn_type"])
            preview = preview.replace("{route}", self.settings.settings["route"])
            
            # Форматируем дату по выбранному формату
            date_format = self.settings.settings.get("date_format", "ГГГГММДД")
            formatted_date = self.format_date_by_format(datetime.now(), date_format)
            preview = preview.replace("{date}", formatted_date)
            
            preview = preview.replace("{counter}", "001")
            preview = preview.replace("{extension}", "jpg")
            preview = preview.replace("{1}", self.settings.settings["var1"])
            preview = preview.replace("{2}", self.settings.settings["var2"])
            preview = preview.replace("{3}", self.settings.settings["var3"])
            
            self.preview_var.set(f"Пример: {preview}")
        except Exception as e:
            self.preview_var.set(f"Ошибка в шаблоне: {e}")
    
    def get_template(self):
        """Получение текущего шаблона"""
        return self.template_text.get("1.0", tk.END).strip()
    
    def apply_template(self):
        """Применение шаблона без закрытия диалога"""
        template = self.get_template()
        if template:
            self.update_preview()
            messagebox.showinfo("Успех", "Шаблон применен для предпросмотра", parent=self.dialog)
    
    def save_template(self):
        """Сохранение шаблона"""
        template = self.get_template()
        if template:
            self.result_template = template
            self.dialog.destroy()

class RenamedFilesManager:
    """Менеджер для хранения информации о переименованных файлах"""
    
    def __init__(self, history_file="renamed_files.json"):
        self.history_file = history_file
        self.renamed_files = set()
        self.load_history()
    
    def load_history(self):
        """Загрузка истории переименований из файла"""
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.renamed_files = set(data.get("renamed_files", []))
                logging.info(f"Загружена история переименований: {len(self.renamed_files)} файлов")
        except Exception as e:
            logging.error(f"Ошибка загрузки истории переименований: {e}")
            self.renamed_files = set()
    
    def save_history(self):
        """Сохранение истории переименований в файл"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump({"renamed_files": list(self.renamed_files)}, f, 
                         ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"Ошибка сохранения истории переименований: {e}")
    
    def add_renamed_file(self, filepath):
        """Добавление файла в историю переименований"""
        file_key = self._get_file_key(filepath)
        self.renamed_files.add(file_key)
        self.save_history()
    
    def is_file_renamed(self, filepath):
        """Проверка, был ли файл уже переименован программой"""
        file_key = self._get_file_key(filepath)
        return file_key in self.renamed_files
    
    def _get_file_key(self, filepath):
        """Создание уникального ключа для файла"""
        try:
            # Используем комбинацию имени файла и времени создания
            stat = os.stat(filepath)
            return f"{Path(filepath).name}_{stat.st_ctime}"
        except:
            return filepath

class Settings:
    """Класс для работы с настройками"""
    def __init__(self, filename="settings.json"):
        self.filename = filename
        self.default_settings = {
            "project": "Проект1",
            "cn_type": "VK",
            "route": "M2.1",
            "number_format": "01",
            "date_format": "ГГГГММДД",
            "var1": "Значение1",
            "var2": "Значение2",
            "var3": "Значение3",
            "folder": r"C:\video\violations",
            "extensions": "png,jpg,jpeg",
            "template": "{project}_{date}_{route}_{counter}_{CN}",
            "monitoring_enabled": True,
            "rename_only_today": True,
            "folder_history": [
                r"C:\video\violations",
                r"C:\temp\files",
                r"D:\projects\images"
            ],
            "template_history": [
                "{project}_{date}_{route}_{counter}_{CN}",
                "{project}_{CN}_{date}_{counter}",
                "{route}_{date}_{counter}_{project}"
            ],
            "enabled_plugins": ["example_plugin"],
            "combobox_values": {
                "project": ["Проект1", "Проект2"],
                "cn_type": ["VK", "Другой"],
                "route": ["M2.1", "M2.2", "M2.3"],
                "number_format": ["1", "01", "001"],
                "date_format": ["ДДММГГГГ", "ДДММГГ", "ГГГГММДД", "ДД.ММ.ГГГГ", "ДД.ММ.ГГ", "ГГГГ.ММ.ДД"],
                "var1": ["Значение1", "Значение2"],
                "var2": ["Значение1", "Значение2"],
                "var3": ["Значение1", "Значение2"]
            },
            "report_route_history": [],
            "column_order": ["number", "create_time", "route", "original_name", "new_name"],
            "column_visibility": {
                "number": True,
                "create_time": True,
                "route": True,
                "original_name": True,
                "new_name": True
            }
        }
        self.load_settings()
    
    def load_settings(self):
        """Загрузка настроек из файла"""
        try:
            if os.path.exists(self.filename):
                with open(self.filename, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    # Объединяем с настройками по умолчанию для совместимости
                    # Обеспечиваем обратную совместимость: если есть старый tl_type, копируем его в cn_type
                    if "tl_type" in loaded_settings and "cn_type" not in loaded_settings:
                        loaded_settings["cn_type"] = loaded_settings["tl_type"]
                    
                    # Обновляем шаблоны в истории для обратной совместимости
                    if "template_history" in loaded_settings:
                        for i, template in enumerate(loaded_settings["template_history"]):
                            if "{TL}" in template:
                                loaded_settings["template_history"][i] = template.replace("{TL}", "{CN}")
                    
                    # Обновляем основной шаблон для обратной совместимости
                    if "template" in loaded_settings and "{TL}" in loaded_settings["template"]:
                        loaded_settings["template"] = loaded_settings["template"].replace("{TL}", "{CN}")
                    
                    # Добавляем новые поля если их нет
                    if "report_route_history" not in loaded_settings:
                        loaded_settings["report_route_history"] = []
                    
                    if "column_order" not in loaded_settings:
                        loaded_settings["column_order"] = self.default_settings["column_order"]
                    
                    if "column_visibility" not in loaded_settings:
                        loaded_settings["column_visibility"] = self.default_settings["column_visibility"]
                    
                    self.settings = {**self.default_settings, **loaded_settings}
            else:
                self.settings = self.default_settings
                self.save_settings()
        except Exception as e:
            logging.error(f"Ошибка загрузки настроек: {e}")
            self.settings = self.default_settings
    
    def save_settings(self):
        """Сохранение настроек в файл"""
        try:
            with open(self.filename, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logging.error(f"Ошибка сохранения настроек: {e}")
    
    def update_setting(self, key, value):
        """Обновление значения настройки"""
        self.settings[key] = value
        self.save_settings()
    
    def add_to_folder_history(self, folder):
        """Добавление папки в историю"""
        if folder and folder not in self.settings["folder_history"]:
            self.settings["folder_history"].insert(0, folder)
            # Ограничиваем историю 10 элементами
            self.settings["folder_history"] = self.settings["folder_history"][:10]
            self.save_settings()
    
    def add_to_template_history(self, template):
        """Добавление шаблона в историю"""
        if template and template not in self.settings["template_history"]:
            self.settings["template_history"].insert(0, template)
            # Ограничиваем историю 10 элементами
            self.settings["template_history"] = self.settings["template_history"][:10]
            self.save_settings()
    
    def add_to_route_history(self, route):
        """Добавление маршрута в историю для отчета"""
        if route and route not in self.settings["report_route_history"]:
            self.settings["report_route_history"].append(route)
            self.save_settings()
    
    def add_to_combobox_values(self, key, value):
        """Добавление значения в список значений комбобокса"""
        if key in self.settings["combobox_values"]:
            if value and value not in self.settings["combobox_values"][key]:
                self.settings["combobox_values"][key].append(value)
                self.save_settings()

class BasePlugin:
    """Базовый класс для всех плагинов"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
    
    def get_tab_name(self):
        """Возвращает название вкладки (должен быть переопределен)"""
        return "Без названия"
    
    def create_tab(self):
        """Создает содержимое вкладки (должен быть переопределен)"""
        return None

class PluginManager:
    """Менеджер плагинов для загрузки дополнительных вкладки"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.plugins = {}
        self.plugin_tabs = {}
    
    def load_plugins(self):
        """Загрузка всех активных плагинов"""
        plugins_dir = "plugins"
        if not os.path.exists(plugins_dir):
            os.makedirs(plugins_dir)
            logging.info(f"Создана папка для плагинов: {plugins_dir}")
            return
        
        enabled_plugins = self.settings.settings.get("enabled_plugins", [])
        
        for plugin_name in enabled_plugins:
            try:
                plugin_path = os.path.join(plugins_dir, f"{plugin_name}.py")
                if os.path.exists(plugin_path):
                    # Динамически импортируем модуль плагина
                    import importlib.util
                    spec = importlib.util.spec_from_file_location(plugin_name, plugin_path)
                    if spec is None:
                        logging.error(f"Не удалось создать spec для плагина: {plugin_name}")
                        continue
                    
                    plugin_module = importlib.util.module_from_spec(spec)
                    
                    try:
                        spec.loader.exec_module(plugin_module)
                    except Exception as e:
                        logging.error(f"Ошибка выполнения модуля плагина {plugin_name}: {e}")
                        continue
                    
                    # Ищем класс плагина (должен наследоваться от BasePlugin)
                    plugin_class = None
                    for name, obj in inspect.getmembers(plugin_module):
                        if (inspect.isclass(obj) and 
                            obj != BasePlugin):
                            # Проверяем, является ли класс плагином (по имени или наследованию)
                            if hasattr(obj, 'get_tab_name') and hasattr(obj, 'create_tab'):
                                plugin_class = obj
                                break
                    
                    if plugin_class:
                        plugin_instance = plugin_class(self.settings, self.root)
                        self.plugins[plugin_name] = plugin_instance
                        logging.info(f"Плагин загружен: {plugin_name}")
                    else:
                        logging.warning(f"Не найден класс плагина в файле: {plugin_name}")
                else:
                    logging.warning(f"Файл плагина не найден: {plugin_path}")
            except Exception as e:
                logging.error(f"Ошибка загрузки плагина {plugin_name}: {e}")
    
    def create_plugin_tabs(self, notebook):
        """Создание вкладок для всех загруженных плагинов"""
        for plugin_name, plugin in self.plugins.items():
            try:
                tab_frame = plugin.create_tab()
                if tab_frame:
                    notebook.add(tab_frame, text=plugin.get_tab_name())
                    self.plugin_tabs[plugin_name] = tab_frame
                    logging.info(f"Создана вкладка для плагина: {plugin_name}")
            except Exception as e:
                logging.error(f"Ошибка создания вкладки для плагина {plugin_name}: {e}")

class FileMonitor:
    """Класс для мониторинга файлов"""
    def __init__(self, settings, rename_callback):
        self.settings = settings
        self.rename_callback = rename_callback
        self.observer = None
        self.event_handler = None
        self.is_monitoring = False
    
    def start_monitoring(self):
        """Запуск мониторинга"""
        if self.is_monitoring:
            return
        
        folder = self.settings.settings["folder"]
        
        if not os.path.exists(folder):
            logging.error(f"Папка не существует: {folder}")
            return False
        
        try:
            self.event_handler = FileHandler(self.settings, self.rename_callback)
            self.observer = Observer()
            self.observer.schedule(self.event_handler, folder, recursive=False)
            self.observer.start()
            self.is_monitoring = True
            logging.info(f"Мониторинг запущен: {folder}")
            return True
        except Exception as e:
            logging.error(f"Ошибка запуска мониторинга: {e}")
            return False
    
    def stop_monitoring(self):
        """Остановка мониторинга"""
        if self.observer and self.is_monitoring:
            try:
                self.observer.stop()
                self.observer.join()
                self.is_monitoring = False
                logging.info("Мониторинг остановлен")
            except Exception as e:
                logging.error(f"Ошибка остановки мониторинга: {e}")

class FileHandler(FileSystemEventHandler):
    """Обработчик событий файловой системы"""
    def __init__(self, settings, rename_callback):
        self.settings = settings
        self.rename_callback = rename_callback
    
    def on_created(self, event):
        """Обработка создания файла с проверкой расширения"""
        if not event.is_directory:
            # Проверяем, включен ли мониторинг
            if self.settings.settings.get("monitoring_enabled", True):
                # Проверяем расширение файла
                file_ext = Path(event.src_path).suffix.lower().lstrip('.')
                extensions = [ext.strip().lower() for ext in self.settings.settings["extensions"].split(",")]
                
                if file_ext in extensions:
                    # Добавляем небольшую задержку для гарантии, что файл полностью создан
                    threading.Timer(1.0, lambda: self.safe_rename_callback(event.src_path)).start()
                else:
                    logging.info(f"Файл {event.src_path} пропущен - расширение {file_ext} не в списке разрешенных")
    
    def safe_rename_callback(self, filepath):
        """Безопасный вызов callback с обработкой исключений"""
        try:
            self.rename_callback([filepath])
        except Exception as e:
            logging.error(f"Критическая ошибка в обработчике переименования: {e}")
            # Не падаем, а просто логируем ошибку

class RenamerApp:
    """Главное приложение"""
    
    def __init__(self, root):
        self.root = root
        self.root.title(f"EGOK Renamer v{VERSION}")
        self.root.geometry("1200x800")
        
        # Установка иконки приложения
        self.set_app_icon()
        
        # Добавляем информацию о разработчике в заголовок
        self.developer_info = "Разработчик: @xDream_Master"
        
        self.settings = Settings()
        self.renamed_files_manager = RenamedFilesManager()
        self.db_manager = DatabaseManager()
        self.monitor = None
        self.log_queue = queue.Queue()
        self.widgets = {}
        self.log_line_counter = 0
        self.rename_history = []
        self.current_route_filter = "Все"
        self.current_date_filter = None
        
        # Данные для отчета
        self.report_data = []
        self.filtered_report_data = []
        
        # Заголовки колонок
        self.column_headers = ["№", "Время создания", "Маршрут", "Исходное имя файла", "Новое имя файла"]
        self.column_ids = ["number", "create_time", "route", "original_name", "new_name"]
        
        # Словарь для управления видимостью колонок
        self.column_visibility = self.settings.settings["column_visibility"]
        
        # Порядок колонок
        self.column_order = self.settings.settings["column_order"]
        
        # Блокировка для безопасного доступа к общим ресурсам
        self.rename_lock = threading.Lock()
        
        # Инициализация менеджера плагинов
        self.plugin_manager = PluginManager(self.settings, self.root)
        
        self.setup_logging()
        self.create_widgets()
        self.load_settings_to_ui()
        self.process_log_queue()
        
        # Загрузка истории переименований из базы данных
        self.load_report_history()
        
        # Запуск мониторинга если включен
        if self.settings.settings.get("monitoring_enabled", True):
            self.start_monitoring()
        
        # Обновляем состояние кнопки и вкладки после инициализации
        self.update_monitoring_button()
    
    def set_app_icon(self):
        """Установка иконки приложения"""
        try:
            # Сначала ищем в текущей директории
            current_dir = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(current_dir, "icon.ico")
            
            # Если не найдено, проверяем MEIPASS (для собранного exe)
            if not os.path.exists(icon_path) and hasattr(sys, '_MEIPASS'):
                icon_path = os.path.join(sys._MEIPASS, "icon.ico")
            
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
                logging.info("Иконка приложения загружена успешно")
            else:
                logging.warning("Файл иконки icon.ico не найден")
        except Exception as e:
            logging.error(f"Ошибка загрузки иконки: {e}")
    
    def setup_logging(self):
        """Настройка логирования в txt файл и GUI"""
        # Создаем папку для логов если нет
        log_dir = "logs"
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # Создаем имя файла с датой
        log_filename = os.path.join(log_dir, f"renamer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
        
        # Настраиваем корневой логгер
        self.logger = logging.getLogger()
        self.logger.setLevel(logging.INFO)
        
        # Очищаем существующие обработчики
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
        
        # Создаем форматтер
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S')
        
        # Обработчик для файла
        file_handler = logging.FileHandler(log_filename, encoding='utf-8')
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        
        # Обработчик для консоли
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        
        # Обработчик для GUI (добавляем в очередь)
        queue_handler = QueueHandler(self.log_queue)
        queue_handler.setLevel(logging.INFO)
        queue_handler.setFormatter(formatter)
        
        # Добавляем обработчики
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        self.logger.addHandler(queue_handler)
        
        logging.info("=" * 50)
        logging.info("ПРОГРАММА ЗАПУЩЕНА")
        logging.info("=" * 50)
        logging.info(f"Логирование инициализировано: {log_filename}")
    
    def create_widgets(self):
        """Создание интерфейса"""
        # Главный фрейм
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Верхняя часть с логотипом
        self.create_header(main_frame)
        
        # Нотебук с вкладками
        self.create_notebook(main_frame)
        
        # Нижняя часть с кнопками
        self.create_footer(main_frame)
    
    def create_header(self, parent):
        """Создание верхней части с логотипом"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Заголовок с информацией о разработчике
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        title_label = ttk.Label(title_frame, text=f"EGOK Renamer v{VERSION}", font=('Arial', 16, 'bold'))
        title_label.pack(anchor=tk.W)
        
        developer_label = ttk.Label(title_frame, text=self.developer_info, font=('Arial', 8), foreground="gray")
        developer_label.pack(anchor=tk.W)
        
        # Логотип - улучшенная обработка путей
        try:
            # Сначала ищем в текущей директории
            current_dir = os.path.dirname(os.path.abspath(__file__))
            image_path = os.path.join(current_dir, "background.png")
            
            # Если не найдено, проверяем MEIPASS (для собранного exe)
            if not os.path.exists(image_path) and hasattr(sys, '_MEIPASS'):
                image_path = os.path.join(sys._MEIPASS, "background.png")
            
            if os.path.exists(image_path):
                image = Image.open(image_path)
                image = image.resize((180, 45), Image.Resampling.LANCZOS)
                self.logo = ImageTk.PhotoImage(image)
                logo_label = ttk.Label(header_frame, image=self.logo)
                logo_label.pack(side=tk.RIGHT)
                logging.info("Логотип загружен успешно")
            else:
                # Заглушка если изображение не найдено
                logo_label = ttk.Label(header_frame, text="[Логотип]", width=20)
                logo_label.pack(side=tk.RIGHT)
                logging.warning("Файл логотипа background.png не найден")
                
        except Exception as e:
            logging.error(f"Ошибка загрузки логотипа: {e}")
            # Заглушка при ошибке
            logo_label = ttk.Label(header_frame, text="[Лого]", width=10)
            logo_label.pack(side=tk.RIGHT)
    
    def create_notebook(self, parent):
        """Создание вкладок"""
        notebook_frame = ttk.Frame(parent)
        notebook_frame.pack(fill=tk.BOTH, expand=True)
        
        # Нотебук для основных функции и плагинов
        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Основная вкладка "ЭГОК" с настройками и логами
        self.egok_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.egok_tab, text="ЭГОК")
        
        # Создаем содержимое вкладки ЭГОК с новой структурой
        self.create_egok_tab(self.egok_tab)
        
        # Загружаем и создаем вкладки плагинов
        self.plugin_manager.load_plugins()
        self.plugin_manager.create_plugin_tabs(self.notebook)
    
    def create_egok_tab(self, parent):
        """Создание основной вкладки ЭГОК с новой структурой"""
        # Создаем PanedWindow для разделения на левую и правую часть
        paned_window = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        # Левая часть - настройки и логи (вертикальное разделение)
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=1)
        
        # Правая часть - отчет (занимает всю правую часть)
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)
        
        # Создаем содержимое левой части (настройки сверху, логи снизу)
        self.create_settings_and_logs_tab(left_frame)
        
        # Создаем содержимое правой части (отчет)
        self.create_report_tab(right_frame)
        
        # Устанавливаем начальное соотношение размеров (40% настройки+логи, 60% отчет)
        paned_window.sashpos(0, int(parent.winfo_reqwidth() * 0.4))
    
    def create_settings_and_logs_tab(self, parent):
        """Создание левой части с настройки и логами"""
        # Создаем вертикальный PanedWindow для разделения настроек и логов
        left_paned = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        left_paned.pack(fill=tk.BOTH, expand=True)
        
        # Верхняя часть - основные настройки
        settings_frame = ttk.LabelFrame(left_paned, text="Основные настройки")
        left_paned.add(settings_frame, weight=1)
        
        # Нижняя часть - логи
        log_frame = ttk.LabelFrame(left_paned, text="Логи программы")
        left_paned.add(log_frame, weight=1)
        
        # Заполняем фреймы
        self.create_settings_content(settings_frame)
        self.create_log_content(log_frame)
        
        # Устанавливаем соотношение (60% настройки, 40% логи)
        left_paned.sashpos(0, int(parent.winfo_reqheight() * 0.6))
    
    def create_settings_content(self, parent):
        """Создание содержимого основных настроек"""
        # Создаем скроллируемую область для настроек
        settings_canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=settings_canvas.yview)
        scrollable_frame = ttk.Frame(settings_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))
        )
        
        settings_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        settings_canvas.configure(yscrollcommand=scrollbar.set)
        
        settings_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Кнопки управления - ПЕРЕМЕЩЕНЫ ВВЕРХ
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10, fill=tk.X)
        
        ttk.Button(button_frame, text="Сохранить настройки", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Управление плагинами", command=self.show_plugins_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Установить плагин", command=self.install_plugin_dialog).pack(side=tk.LEFT, padx=5)
        
        # КНОПКА МОНИТОРИНГА ПЕРЕМЕЩЕНА СЮДА - В САМЫЙ ВЕРХ
        monitoring_button_frame = ttk.Frame(button_frame)
        monitoring_button_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(monitoring_button_frame, text="Мониторинг:").pack(side=tk.LEFT)
        
        # Создаем кастомный стиль для кнопки-переключателя
        style = ttk.Style()
        style.configure("Green.TButton", background="#4CAF50", foreground="#4CAF50")
        style.configure("Red.TButton", background="#F44336", foreground="#F44336")
        
        self.monitoring_button = ttk.Button(
            monitoring_button_frame, 
            text="ВКЛ", 
            style="Green.TButton",
            command=self.toggle_monitoring,
            width=8
        )
        self.monitoring_button.pack(side=tk.LEFT, padx=5)
        
        # Обновляем состояние кнопки при запуске
        self.update_monitoring_button()
        
        self.create_combobox_row(scrollable_frame, "Проект:", "project", 0)
        self.create_combobox_row(scrollable_frame, "Тип ЦН:", "cn_type", 1)
        self.create_combobox_row(scrollable_frame, "Маршрут:", "route", 2)
        self.create_combobox_row(scrollable_frame, "Формат номера:", "number_format", 3)
        self.create_combobox_row(scrollable_frame, "Формат даты:", "date_format", 4)
        self.create_combobox_row(scrollable_frame, "Переменная 1:", "var1", 5)
        self.create_combobox_row(scrollable_frame, "Переменная 2:", "var2", 6)
        self.create_combobox_row(scrollable_frame, "Переменная 3:", "var3", 7)
        
        # Настройки переименования
        rename_frame = ttk.LabelFrame(scrollable_frame, text="Настройки переименования")
        rename_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Папка с файлами
        folder_frame = ttk.Frame(rename_frame)
        folder_frame.pack(fill=tk.X, pady=2)
        ttk.Label(folder_frame, text="Папка с файлами:").pack(side=tk.LEFT)
        
        folder_var = tk.StringVar(value=self.settings.settings["folder"])
        self.widgets["folder_var"] = folder_var
        
        # Combobox для выбора папки из истории
        folder_cb = ttk.Combobox(
            folder_frame, 
            textvariable=folder_var, 
            values=self.settings.settings.get("folder_history", []),
            width=30
        )
        folder_cb.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Привязываем события для обновления истории
        folder_cb.bind('<<ComboboxSelected>>', lambda e: self.on_folder_selected())
        folder_cb.bind('<FocusOut>', lambda e: self.on_folder_selected())
        
        ttk.Button(folder_frame, text="Обзор", command=self.browse_folder).pack(side=tk.LEFT)
        
        # Расширения файлов
        ext_frame = ttk.Frame(rename_frame)
        ext_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ext_frame, text="Расширения файлов:").pack(side=tk.LEFT)
        
        ext_var = tk.StringVar(value=self.settings.settings["extensions"])
        self.widgets["ext_var"] = ext_var
        
        ext_entry = ttk.Entry(ext_frame, textvariable=ext_var, width=30)
        ext_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Шаблон имени - УЛУЧШЕННАЯ ВЕРСИЯ
        template_frame = ttk.Frame(rename_frame)
        template_frame.pack(fill=tk.X, pady=2)
        ttk.Label(template_frame, text="Шаблон имени:").pack(side=tk.LEFT)
        
        template_var = tk.StringVar(value=self.settings.settings["template"])
        self.widgets["template_var"] = template_var
        
        # Combobox для выбора шаблона из истории
        self.template_cb = ttk.Combobox(
            template_frame, 
            textvariable=template_var, 
            values=self.settings.settings.get("template_history", []),
            width=30
        )
        self.template_cb.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Привязываем события для обновления истории
        self.template_cb.bind('<<ComboboxSelected>>', lambda e: self.on_template_selected())
        self.template_cb.bind('<FocusOut>', lambda e: self.on_template_selected())
        
        # НОВАЯ КНОПКА - КОНСТРУКТОР ШАБЛОНА
        ttk.Button(template_frame, text="Конструктор", 
                  command=self.open_template_builder).pack(side=tk.LEFT, padx=2)
        ttk.Button(template_frame, text="Проверить", 
                  command=self.check_template).pack(side=tk.LEFT, padx=2)
        
        # Опция переименовывать только сегодняшние файлы
        today_only_frame = ttk.Frame(rename_frame)
        today_only_frame.pack(fill=tk.X, pady=2)
        
        self.rename_only_today_var = tk.BooleanVar(value=self.settings.settings.get("rename_only_today", True))
        self.widgets["rename_only_today_var"] = self.rename_only_today_var
        
        rename_only_today_cb = ttk.Checkbutton(
            today_only_frame, 
            text="Переименовывать только сегодняшние файлы",
            variable=self.rename_only_today_var
        )
        rename_only_today_cb.pack(anchor=tk.W)
    
    def open_template_builder(self):
        """Открытие конструктора шаблонов"""
        current_template = self.widgets["template_var"].get()
        
        dialog = TemplateBuilderDialog(self.root, current_template, self.settings)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result_template:
            # Обновляем поле шаблона
            self.widgets["template_var"].set(dialog.result_template)
            
            # Сохраняем в настройки
            self.settings.update_setting("template", dialog.result_template)
            self.settings.add_to_template_history(dialog.result_template)
            
            # Обновляем список в комбобоксе
            self.template_cb['values'] = self.settings.settings.get("template_history", [])
            
            logging.info(f"Шаблон обновлен через конструктор: {dialog.result_template}")
            
            # Показываем предпросмотр нового шаблона
            self.check_template()
    
    def create_log_content(self, parent):
        """Создание содержимого логов"""
        # Текстовое поле для логов с поддержкой цветов
        self.log_text = tk.Text(parent, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.tag_configure("black", foreground="black")
        self.log_text.tag_configure("gray", foreground="gray")
        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("warning", foreground="orange")
        self.log_text.tag_configure("info", foreground="blue")
        
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def create_report_tab(self, parent):
        """Создание правой части с отчетом"""
        # Отчет о переименованных файлах
        report_frame = ttk.LabelFrame(parent, text="Отчет о переименованных файлах")
        report_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Фрейм для кнопок управления отчетом и фильтра (над таблицей)
        report_controls_frame = ttk.Frame(report_frame)
        report_controls_frame.pack(fill=tk.X, pady=5)
        
        # Кнопки управления отчетом
        ttk.Button(report_controls_frame, text="Копировать выделенное", command=self.copy_selected_cells).pack(side=tk.LEFT, padx=2)
        ttk.Button(report_controls_frame, text="Копировать все", command=self.copy_all_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(report_controls_frame, text="Очистить отчет", command=self.clear_report).pack(side=tk.LEFT, padx=2)
        ttk.Button(report_controls_frame, text="Экспорт в файл", command=self.export_report).pack(side=tk.LEFT, padx=2)
        
        # Кнопка управления колонками
        ttk.Button(report_controls_frame, text="Управление колонками", command=self.show_column_management_dialog).pack(side=tk.LEFT, padx=2)
        
        # Фильтр по маршруту и дате
        filter_frame = ttk.Frame(report_controls_frame)
        filter_frame.pack(side=tk.RIGHT, padx=5)
        
        # Фильтр по маршруту
        route_filter_frame = ttk.Frame(filter_frame)
        route_filter_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(route_filter_frame, text="Маршрут:").pack(side=tk.LEFT, padx=2)
        
        self.route_filter_var = tk.StringVar(value="Все")
        
        # Получаем значения для фильтра из истории маршрутов отчета
        route_values = ["Все"] + self.settings.settings.get("report_route_history", [])
        
        self.route_filter_cb = ttk.Combobox(
            route_filter_frame, 
            textvariable=self.route_filter_var,
            values=route_values,
            state="readonly",
            width=10
        )
        self.route_filter_cb.pack(side=tk.LEFT, padx=2)
        self.route_filter_cb.bind('<<ComboboxSelected>>', self.on_route_filter_changed)
        
        # Фильтр по дате
        date_filter_frame = ttk.Frame(filter_frame)
        date_filter_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(date_filter_frame, text="Дата:").pack(side=tk.LEFT, padx=2)
        
        # Получаем все доступные даты из базы данных
        available_dates = self.db_manager.get_all_dates()
        date_values = ["Все даты"] + available_dates
        
        self.date_filter_var = tk.StringVar(value="Все даты")
        
        self.date_filter_cb = ttk.Combobox(
            date_filter_frame, 
            textvariable=self.date_filter_var,
            values=date_values,
            state="readonly",
            width=12
        )
        self.date_filter_cb.pack(side=tk.LEFT, padx=2)
        self.date_filter_cb.bind('<<ComboboxSelected>>', self.on_date_filter_changed)
        
        # Кнопка обновления списка дат
        ttk.Button(date_filter_frame, text="Обновить", command=self.update_date_filter).pack(side=tk.LEFT, padx=2)
        
        # Фрейм для таблицы отчета
        table_frame = ttk.Frame(report_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Создаем улучшенную таблицу с поддержкой выделения ячеек
        if TKSHEET_AVAILABLE:
            self.create_sheet_table(table_frame)
        else:
            self.create_fallback_table(table_frame)
        
        # Применяем настройки видимости колонок
        self.apply_column_visibility()
    
    def create_sheet_table(self, parent):
        """Создание продвинутой таблицы с улучшенным выделением как в Excel"""
        try:
            # Создаем таблицу с включенными индексами строк
            self.report_sheet = tksheet.Sheet(
                parent,
                show_row_index=True,  # Включаем отображение индексов строк
                show_header=True,
                show_x_scrollbar=True,
                show_y_scrollbar=True,
                headers=self.column_headers,
                header_height=30,
                row_index_width=50,
                frame_bg="white"
            )
            
            # Включаем все необходимые привязки для Excel-подобного поведения
            self.report_sheet.enable_bindings(
                # Основное выделение
                "single_select",
                "toggle_select",
                "drag_select",
                "row_select",
                "column_select",
                "cell_select",
                "all",
                
                # Навигация
                "arrowkeys",
                "tab",
                "ctrl_a",
                "ctrl_c",
                "ctrl_v",
                "ctrl_x",
                
                # Работа с буфером обмена
                "copy",
                "cut",
                "paste",
                "delete",
                
                # Редактирование
                "edit_cell",
                
                # Контекстное меню
                "right_click_popup_menu",
                "rc_select",
                "rc_insert_column",
                "rc_delete_column",
                "rc_insert_row",
                "rc_delete_row",
                
                # Дополнительные функции
                "undo",
                "redo",
                "edit_header",
                
                # Перетаскивание колонок
                "drag_and_drop_column"
            )
            
            # Настраиваем таблицу для лучшего отображения
            self.report_sheet.set_sheet_data(self.report_data)
            
            # Настраиваем ширину колонок
            self.report_sheet.column_width(column=0, width=50)   # №
            self.report_sheet.column_width(column=1, width=120)  # Время создания
            self.report_sheet.column_width(column=2, width=100)  # Маршрут
            self.report_sheet.column_width(column=3, width=250)  # Исходное имя файла
            self.report_sheet.column_width(column=4, width=250)  # Новое имя файла
            
            # Настраиваем выравнивание
            if self.report_data:
                self.report_sheet.set_cell_alignments(align="center", cells=[(r, 0) for r in range(len(self.report_data))])  # № по центру
                self.report_sheet.set_cell_alignments(align="center", cells=[(r, 1) for r in range(len(self.report_data))])  # Время по центру
                self.report_sheet.set_cell_alignments(align="center", cells=[(r, 2) for r in range(len(self.report_data))])  # Маршрут по центру
            
            # Упаковка таблицы
            self.report_sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Создаем улучшенное контекстное меню
            self.create_enhanced_context_menu()
            
            # Привязываем улучшенное контекстное меню
            self.report_sheet.bind("<Button-3>", self.show_enhanced_context_menu)
            
            # Привязываем двойной клик для редактирования
            self.report_sheet.bind("<Double-Button-1>", self.on_double_click)
            
            logging.info("Таблица отчета инициализирована с улучшенным выделением")
            
        except Exception as e:
            logging.error(f"Ошибка создания таблицы tksheet: {e}")
            self.create_fallback_table(parent)
    
    def create_enhanced_context_menu(self):
        """Создание улучшенного контекстного меню для таблицы"""
        self.enhanced_context_menu = tk.Menu(self.report_sheet, tearoff=0)
        
        # Основные операции с выделением
        self.enhanced_context_menu.add_command(
            label="Копировать выделенное", 
            command=self.copy_selected_cells
        )
        self.enhanced_context_menu.add_command(
            label="Копировать как текст", 
            command=self.copy_as_text
        )
        self.enhanced_context_menu.add_separator()
        
        # Опции выделения
        selection_menu = tk.Menu(self.enhanced_context_menu, tearoff=0)
        selection_menu.add_command(label="Выделить всю строку", command=self.select_current_row)
        selection_menu.add_command(label="Выделить весь столбец", command=self.select_current_column)
        selection_menu.add_command(label="Выделить всю таблицу", command=self.select_all_table)
        self.enhanced_context_menu.add_cascade(label="Выделить", menu=selection_menu)
        
        self.enhanced_context_menu.add_separator()
        
        # Управление столбцами
        columns_menu = tk.Menu(self.enhanced_context_menu, tearoff=0)
        columns_menu.add_command(label="Скрыть столбец", command=self.hide_current_column)
        columns_menu.add_command(label="Показать все столбцы", command=self.show_all_columns)
        columns_menu.add_separator()
        columns_menu.add_command(label="Настроить столбцы...", command=self.show_column_management_dialog)
        self.enhanced_context_menu.add_cascade(label="Столбцы", menu=columns_menu)
        
        self.enhanced_context_menu.add_separator()
        
        # Дополнительные функции
        self.enhanced_context_menu.add_command(label="Экспорт выделенного", command=self.export_selected)
        self.enhanced_context_menu.add_command(label="Очистить отчет", command=self.clear_report)
        
        logging.info("Улучшенное контекстное меню создано")
    
    def create_fallback_table(self, parent):
        """Создание резервной таблицы с помощью Treeview (если tksheet не доступен)"""
        # Создаем Treeview для отчета с новыми колонками
        columns = ("number", "create_time", "route", "original_name", "new_name")
        self.report_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode='extended')
        
        # Настраиваем заголовки колонок
        self.report_tree.heading("number", text="№")
        self.report_tree.heading("create_time", text="Время")
        self.report_tree.heading("route", text="Маршрут")
        self.report_tree.heading("original_name", text="Исходное имя")
        self.report_tree.heading("new_name", text="Новое имя")
        
        # Настраиваем колонки для максимального расширения
        self.report_tree.column("number", width=40, minwidth=40, stretch=False)
        self.report_tree.column("create_time", width=120, minwidth=120, stretch=False)
        self.report_tree.column("route", width=80, minwidth=80, stretch=False)
        self.report_tree.column("original_name", width=200, minwidth=150, stretch=True)
        self.report_tree.column("new_name", width=200, minwidth=150, stretch=True)
        
        # Scrollbar для Treeview
        report_scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=report_scrollbar.set)
        
        # Упаковка таблицы
        self.report_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        report_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Контекстное меню для копирования
        self.report_context_menu = tk.Menu(self.report_tree, tearoff=0)
        self.report_context_menu.add_command(label="Копировать выделенные строки", command=self.copy_selected_files)
        self.report_context_menu.add_command(label="Копировать всю таблицу", command=self.copy_all_files)
        self.report_context_menu.add_separator()
        self.report_context_menu.add_command(label="Очистить отчет", command=self.clear_report)
        
        # Привязываем контекстное меню
        self.report_tree.bind("<Button-3>", self.show_report_context_menu)
    
    def show_enhanced_context_menu(self, event):
        """Показать улучшенное контекстное меню"""
        try:
            self.enhanced_context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            logging.error(f"Ошибка показа контекстного меню: {e}")
    
    def on_double_click(self, event):
        """Обработка двойного клика для редактирования ячейки"""
        # tksheet автоматически обрабатывает редактирование при включенных binding
        pass
    
    def copy_as_text(self):
        """Копирование выделенного как форматированный текст"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if not selected:
                messagebox.showinfo("Информация", "Не выделены ячейки для копирования")
                return
            
            # Получаем данные выделенных ячеек
            text_lines = []
            current_row = None
            current_line = []
            
            for row, col in selected:
                if current_row != row and current_line:
                    text_lines.append("\t".join(current_line))
                    current_line = []
                current_row = row
                
                cell_value = self.report_sheet.get_cell_data(row, col)
                current_line.append(str(cell_value if cell_value is not None else ""))
            
            if current_line:
                text_lines.append("\t".join(current_line))
            
            # Копируем в буфер обмена
            text_to_copy = "\n".join(text_lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(text_to_copy)
            
            messagebox.showinfo("Успех", f"Скопировано {len(selected)} ячеек как текст")
            
        except Exception as e:
            logging.error(f"Ошибка копирования как текст: {e}")
            messagebox.showerror("Ошибка", f"Не удалось скопировать данные: {e}")
    
    def select_current_row(self):
        """Выделить текущую строку"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if selected:
                row = selected[0][0]
                self.report_sheet.select_row(row)
        except Exception as e:
            logging.error(f"Ошибка выделения строки: {e}")
    
    def select_current_column(self):
        """Выделить текущий столбец"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if selected:
                col = selected[0][1]
                self.report_sheet.select_column(col)
        except Exception as e:
            logging.error(f"Ошибка выделения столбца: {e}")
    
    def select_all_table(self):
        """Выделить всю таблицу"""
        try:
            self.report_sheet.select_all()
        except Exception as e:
            logging.error(f"Ошибка выделения всей таблицы: {e}")
    
    def hide_current_column(self):
        """Скрыть текущий столбец"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if selected:
                col = selected[0][1]
                self.report_sheet.hide_columns(columns=[col])
                messagebox.showinfo("Успех", f"Столбец {col} скрыт")
        except Exception as e:
            logging.error(f"Ошибка скрытия столбца: {e}")
    
    def show_all_columns(self):
        """Показать все скрытые столбцы"""
        try:
            self.report_sheet.display_columns("all")
            messagebox.showinfo("Успех", "Все столбцы показаны")
        except Exception as e:
            logging.error(f"Ошибка показа всех столбцов: {e}")
    
    def export_selected(self):
        """Экспорт выделенных данных в файл"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if not selected:
                messagebox.showwarning("Внимание", "Не выделены данные для экспорта")
                return
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[
                    ("CSV files", "*.csv"),
                    ("Text files", "*.txt"), 
                    ("All files", "*.*")
                ],
                title="Экспорт выделенных данных"
            )
            
            if file_path:
                # Получаем выделенные данные
                data_to_export = []
                
                # Собираем строки с данными
                rows_data = {}
                for row, col in selected:
                    if row not in rows_data:
                        rows_data[row] = {}
                    cell_value = self.report_sheet.get_cell_data(row, col)
                    rows_data[row][col] = cell_value if cell_value is not None else ""
                
                # Сортируем по строкам и столбцам
                for row in sorted(rows_data.keys()):
                    row_data = []
                    for col in sorted(rows_data[row].keys()):
                        row_data.append(str(rows_data[row][col]))
                    data_to_export.append(",".join(row_data))
                
                # Записываем в файл
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("\n".join(data_to_export))
                
                messagebox.showinfo("Успех", f"Данные экспортированы в:\n{file_path}")
                
        except Exception as e:
            logging.error(f"Ошибка экспорта выделенных данных: {e}")
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {e}")
    
    def show_report_context_menu(self, event):
        """Показать контекстное меню для отчета"""
        self.report_context_menu.post(event.x_root, event.y_root)
    
    def copy_selected_cells(self):
        """Копировать выделенные ячейки в буфер обмена с улучшенной обработкой"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            try:
                # Пробуем разные методы копирования в зависимости от версии tksheet
                if hasattr(self.report_sheet, 'ctrl_c'):
                    self.report_sheet.ctrl_c()
                elif hasattr(self.report_sheet, 'copy'):
                    self.report_sheet.copy()
                else:
                    # Альтернативный метод копирования через выделение данных
                    self.copy_selected_cells_manual()
                    return
                
                # Получаем информацию о выделении для пользователя
                selected = self.report_sheet.get_selected_cells()
                if selected:
                    # Группируем по строкам для подсчета
                    rows = set()
                    cols = set()
                    for row, col in selected:
                        rows.add(row)
                        cols.add(col)
                    
                    messagebox.showinfo(
                        "Успех", 
                        f"Скопировано:\n"
                        f"- Ячеек: {len(selected)}\n"
                        f"- Строк: {len(rows)}\n"
                        f"- Столбцов: {len(cols)}\n\n"
                        f"Данные помещены в буфер обмена"
                    )
                else:
                    messagebox.showinfo("Информация", "Не выделены ячейки для копирования")
                    
            except Exception as e:
                logging.error(f"Ошибка копирования ячеек: {e}")
                # Пробуем альтернативный метод
                try:
                    self.copy_selected_cells_manual()
                except Exception as e2:
                    logging.error(f"Ошибка альтернативного копирования: {e2}")
                    messagebox.showerror("Ошибка", f"Не удалось скопировать ячейки: {e}")
        else:
            # Резервный метод для Treeview
            self.copy_selected_files()

    def copy_selected_cells_manual(self):
        """Альтернативный метод копирования выделенных ячеек"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if not selected:
                messagebox.showinfo("Информация", "Не выделены ячейки для копирования")
                return
            
            # Собираем данные из выделенных ячеек
            text_lines = []
            current_row = None
            current_line = []
            
            # Сортируем ячейки по строкам и столбцам
            sorted_cells = sorted(selected, key=lambda x: (x[0], x[1]))
            
            for row, col in sorted_cells:
                if current_row != row and current_line:
                    text_lines.append("\t".join(current_line))
                    current_line = []
                current_row = row
                
                cell_value = self.report_sheet.get_cell_data(row, col)
                current_line.append(str(cell_value if cell_value is not None else ""))
            
            if current_line:
                text_lines.append("\t".join(current_line))
            
            # Копируем в буфер обмена
            text_to_copy = "\n".join(text_lines)
            self.root.clipboard_clear()
            self.root.clipboard_append(text_to_copy)
            
            messagebox.showinfo("Успех", f"Скопировано {len(selected)} ячеек")
            
        except Exception as e:
            logging.error(f"Ошибка ручного копирования ячеек: {e}")
            raise
    
    def copy_selected_files(self):
        """Копировать выделенные строки в буфер обмена (для Treeview)"""
        if not TKSHEET_AVAILABLE and hasattr(self, 'report_tree'):
            selected_items = self.report_tree.selection()
            if not selected_items:
                messagebox.showwarning("Внимание", "Не выделены строки для копирования")
                return
            
            # Собираем все данные выделенных строк
            all_lines = []
            for item in selected_items:
                values = self.report_tree.item(item, "values")
                if values:
                    # Формируем строку с табуляцией между значениями
                    line = "\t".join(str(value) for value in values)
                    all_lines.append(line)
            
            if all_lines:
                # Копируем в буфер обмена
                self.root.clipboard_clear()
                self.root.clipboard_append("\n".join(all_lines))
                messagebox.showinfo("Успех", f"Скопировано {len(all_lines)} строк в буфер обмена")
    
    def copy_all_files(self):
        """Копировать все данные отчета в буфер обмена"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            # Выделяем всю таблицу и копируем
            try:
                self.report_sheet.select_all()
                
                # Пробуем разные методы копирования
                if hasattr(self.report_sheet, 'ctrl_c'):
                    self.report_sheet.ctrl_c()
                elif hasattr(self.report_sheet, 'copy'):
                    self.report_sheet.copy()
                else:
                    # Альтернативный метод копирования всех данных
                    self.copy_all_files_manual()
                    return
                
                self.report_sheet.deselect("all")
                messagebox.showinfo("Успех", "Вся таблица скопирована в буфер обмена")
            except Exception as e:
                logging.error(f"Ошибка копирования таблицы: {e}")
                # Пробуем альтернативный метод
                try:
                    self.copy_all_files_manual()
                except Exception as e2:
                    logging.error(f"Ошибка альтернативного копирования всей таблицы: {e2}")
                    messagebox.showerror("Ошибка", f"Не удалось скопировать таблицу: {e}")
        elif hasattr(self, 'report_tree'):
            # Резервный метод для Treeview
            all_items = self.report_tree.get_children()
            if not all_items:
                messagebox.showwarning("Внимание", "В отчете нет данных")
                return
            
            # Собираем все данные всех строк
            all_lines = []
            for item in all_items:
                values = self.report_tree.item(item, "values")
                if values:
                    # Формируем строку с табуляцией между значениями
                    line = "\t".join(str(value) for value in values)
                    all_lines.append(line)
            
            if all_lines:
                # Копируем в буфер обмена
                self.root.clipboard_clear()
                self.root.clipboard_append("\n".join(all_lines))
                messagebox.showinfo("Успех", f"Скопировано {len(all_lines)} строк в буфер обмена")

    def copy_all_files_manual(self):
        """Альтернативный метод копирования всех данных таблицы"""
        try:
            data = self.report_sheet.get_sheet_data()
            if not data:
                messagebox.showwarning("Внимание", "В отчете нет данных")
                return
            
            text_lines = []
            for row in data:
                if row and any(cell is not None for cell in row):
                    line = "\t".join(str(cell) if cell is not None else "" for cell in row)
                    text_lines.append(line)
            
            if text_lines:
                self.root.clipboard_clear()
                self.root.clipboard_append("\n".join(text_lines))
                messagebox.showinfo("Успех", f"Скопировано {len(text_lines)} строк в буфер обмена")
            
        except Exception as e:
            logging.error(f"Ошибка ручного копирования всей таблицы: {e}")
            raise
    
    def clear_report(self):
        """Очистить отчет"""
        if messagebox.askyesno("Подтверждение", "Очистить отчет о переименованных файлах?"):
            if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
                self.report_sheet.set_sheet_data([])
                self.report_data = []
            elif hasattr(self, 'report_tree'):
                for item in self.report_tree.get_children():
                    self.report_tree.delete(item)
            
            self.rename_history.clear()
            # Сбрасываем фильтры
            self.route_filter_var.set("Все")
            self.date_filter_var.set("Все даты")
            self.current_route_filter = "Все"
            self.current_date_filter = None
            
            # Очищаем базу данных
            self.db_manager.clear_all_records()
            
            logging.info("Отчет о переименованных файлах очищен")
    
    def export_report(self):
        """Экспорт отчета в файл"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            data = self.report_sheet.get_sheet_data()
            if not data or len(data) == 0:
                messagebox.showwarning("Внимание", "В отчете нет данных для экспорта")
                return
        elif hasattr(self, 'report_tree'):
            all_items = self.report_tree.get_children()
            if not all_items:
                messagebox.showwarning("Внимание", "В отчете нет данных для экспорта")
                return
        else:
            messagebox.showwarning("Внимание", "Отчет не инициализирован")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")],
            title="Экспорт отчета"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("Отчет о переименованных файлах\n")
                    f.write("=" * 80 + "\n")
                    f.write(f"Создан: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Фильтр по маршруту: {self.current_route_filter}\n")
                    f.write(f"Фильтр по дате: {self.current_date_filter or 'Все даты'}\n")
                    f.write("=" * 80 + "\n")
                    f.write("№\tВремя\tМаршрут\tИсходное имя\tНовое имя\n")
                    f.write("-" * 80 + "\n")
                    
                    if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
                        # Экспорт данных из tksheet
                        for row in self.report_sheet.get_sheet_data():
                            if row and any(cell is not None for cell in row):
                                f.write(f"{row[0] or ''}\t{row[1] or ''}\t{row[2] or ''}\t{row[3] or ''}\t{row[4] or ''}\n")
                    else:
                        # Экспорт данных из Treeview
                        for item in self.report_tree.get_children():
                            values = self.report_tree.item(item, "values")
                            if values and len(values) > 4:
                                f.write(f"{values[0]}\t{values[1]}\t{values[2]}\t{values[3]}\t{values[4]}\n")
                
                messagebox.showinfo("Успех", f"Отчет экспортирован в файл:\n{file_path}")
                logging.info(f"Отчет экспортирован в файл: {file_path}")
            except Exception as e:
                logging.error(f"Ошибка экспорта отчета: {e}")
                messagebox.showerror("Ошибка", f"Ошибка экспорта отчета: {e}")
    
    def add_to_report(self, original_name, new_name, filepath, create_time=None):
        """Добавление записи в отчет о переименовании"""
        # Если время не передано, пытаемся получить из файла
        if create_time is None:
            try:
                create_time = datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%H:%M:%S')
            except Exception as e:
                logging.error(f"Ошибка получения времени создания файла {filepath}: {e}")
                create_time = datetime.now().strftime('%H:%M:%S')
        
        # Получаем текущий маршрут из настроек
        route = self.settings.settings["route"]
        
        # Добавляем маршрут в историю для фильтра
        if route not in self.settings.settings.get("report_route_history", []):
            self.settings.add_to_route_history(route)
            # Обновляем комбобокс фильтра
            self.update_route_filter_combobox()
        
        # Номер строки
        number = len(self.rename_history) + 1
        
        # Добавляем в историю
        self.rename_history.append({
            "number": number,
            "create_time": create_time,
            "route": route,
            "original_name": original_name,
            "new_name": new_name
        })
        
        # Создаем строку данных
        row_data = [number, create_time, route, original_name, new_name]
        
        # Добавляем в данные отчета
        self.report_data.append(row_data)
        
        # Сохраняем в базу данных
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.db_manager.add_record(timestamp, route, original_name, new_name, filepath)
        
        # Добавляем в таблицу
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            self.report_sheet.set_sheet_data(self.report_data)
        elif hasattr(self, 'report_tree'):
            values = (number, create_time, route, original_name, new_name)
            item_id = self.report_tree.insert("", tk.END, values=values)
            
            # Автоматически прокручиваем к последней записи
            self.report_tree.see(item_id)
        
        # Логируем добавление в отчет
        logging.info(f"Добавлено в отчет: {original_name} -> {new_name}")
    
    def is_record_from_today(self, timestamp):
        """Проверяет, относится ли запись к сегодняшней дате"""
        try:
            record_date = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S").date()
            today = datetime.now().date()
            return record_date == today
        except:
            return False
    
    def load_report_history(self):
        """Загрузка истории переименований из базы данных"""
        try:
            records = self.db_manager.get_records_by_date()
            
            # Преобразуем записи в формат для отчета
            self.report_data = []
            for i, record in enumerate(records):
                row_data = [
                    i + 1,
                    record['timestamp'].split(' ')[1] if ' ' in record['timestamp'] else record['timestamp'],
                    record['route'],
                    record['original_name'],
                    record['new_name']
                ]
                self.report_data.append(row_data)
            
            # Обновляем таблицу
            if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
                self.report_sheet.set_sheet_data(self.report_data)
            elif hasattr(self, 'report_tree'):
                for item in self.report_tree.get_children():
                    self.report_tree.delete(item)
                
                for row in self.report_data:
                    self.report_tree.insert("", tk.END, values=tuple(row))
            
            # Обновляем список дат в фильтре
            self.update_date_filter()
            
            logging.info(f"Загружено {len(records)} записей из истории")
        except Exception as e:
            logging.error(f"Ошибка загрузки истории отчета: {e}")
    
    def update_route_filter_combobox(self):
        """Обновление комбобокса фильтра по маршруту"""
        # Обновляем значения комбобокса фильтра
        route_values = ["Все"] + self.settings.settings.get("report_route_history", [])
        self.route_filter_cb['values'] = route_values
    
    def update_date_filter(self):
        """Обновление комбобокса фильтра по дате"""
        # Получаем все доступные даты из базы данных
        available_dates = self.db_manager.get_all_dates()
        date_values = ["Все даты"] + available_dates
        self.date_filter_cb['values'] = date_values
    
    def apply_column_visibility(self):
        """Применить настройки видимости колонок"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            # Для tksheet используем свойство visible_columns
            columns_to_show = []
            
            # Создаем mapping: column_id -> индекс
            column_id_to_index = {column_id: i for i, column_id in enumerate(self.column_ids)}
            
            # Формируем список индексов видимых колонок в правильном порядке
            for column_id in self.column_order:
                if self.column_visibility.get(column_id, True):
                    if column_id in column_id_to_index:
                        columns_to_show.append(column_id_to_index[column_id])
            
            # Устанавливаем видимые колонки через свойство visible_columns
            try:
                self.report_sheet.visible_columns = columns_to_show
                # ОБНОВЛЯЕМ ДАННЫЕ В ТАБЛИЦЕ - ИСПРАВЛЕНИЕ БАГА
                self.report_sheet.set_sheet_data(self.report_data)
            except AttributeError:
                # Для старых версий tksheet
                logging.warning("Свойство visible_columns недоступно, используется display_columns")
                self.report_sheet.display_columns(columns_to_show)
                # ОБНОВЛЯЕМ ДАННЫЕ В ТАБЛИЦЕ - ИСПРАВЛЕНИЕ БАГА
                self.report_sheet.set_sheet_data(self.report_data)
                
        elif hasattr(self, 'report_tree'):
            # Для Treeview определяем видимые колонки в правильном порядке
            visible_columns = [col for col in self.column_order if self.column_visibility.get(col, True)]
            
            # Устанавливаем отображаемые колонки
            self.report_tree["displaycolumns"] = visible_columns
            
            # Обновляем заголовки для видимых колонок
            column_names = {
                "number": "№",
                "create_time": "Время",
                "route": "Маршрут",
                "original_name": "Исходное имя",
                "new_name": "Новое имя"
            }
            
            for column in visible_columns:
                self.report_tree.heading(column, text=column_names.get(column, column))
    
    def on_route_filter_changed(self, event=None):
        """Обработка изменения фильтра по маршруту"""
        selected_route = self.route_filter_var.get()
        self.current_route_filter = selected_route
        self.apply_filters()
    
    def on_date_filter_changed(self, event=None):
        """Обработка изменения фильтра по дате"""
        selected_date = self.date_filter_var.get()
        self.current_date_filter = selected_date if selected_date != "Все даты" else None
        self.apply_filters()
    
    def apply_filters(self):
        """Применение всех активных фильтров"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            # Для tksheet фильтруем данные
            filtered_data = []
            
            for row in self.report_data:
                route_match = (self.current_route_filter == "Все" or 
                              (len(row) > 2 and row[2] == self.current_route_filter))
                
                # Для фильтрации по дате нам нужно получить полные данные из БД
                # Пока пропускаем фильтрацию по дате для tksheet для простоты
                if route_match:
                    filtered_data.append(row)
            
            self.report_sheet.set_sheet_data(filtered_data)
        elif hasattr(self, 'report_tree'):
            # Для Treeview показываем/скрываем элементы в соответствии с фильтрами
            all_items = self.report_tree.get_children()
            
            for item in all_items:
                values = self.report_tree.item(item, "values")
                if len(values) > 2:
                    route = values[2]  # values[2] - колонка с маршрутом
                    route_match = (self.current_route_filter == "Все" or route == self.current_route_filter)
                    
                    # Пока пропускаем фильтрацию по дате для Treeview для простоты
                    if route_match:
                        # Показываем элемент
                        self.report_tree.attach(item, '', 'end')
                    else:
                        # Скрываем элемент (но не удаляем)
                        self.report_tree.detach(item)
        
        logging.info(f"Применены фильтры: маршрут={self.current_route_filter}, дата={self.current_date_filter or 'Все даты'}")
    
    def create_combobox_row(self, parent, label, key, row):
        """Создание строки с Combobox"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=2, padx=5)
        
        ttk.Label(frame, text=label, width=15).pack(side=tk.LEFT)
        
        values = self.settings.settings["combobox_values"][key]
        current_value = self.settings.settings[key]
        
        var = tk.StringVar(value=current_value)
        self.widgets[f"{key}_var"] = var
        
        combobox = ttk.Combobox(frame, textvariable=var, values=values, width=20)
        combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Подсказка
        hints = {
            "project": "{project}",
            "cn_type": "{CN}",
            "route": "{route}",
            "number_format": "{counter}",
            "date_format": "{date}",
            "var1": "{1}",
            "var2": "{2}",
            "var3": "{3}"
        }
        ttk.Label(frame, text=f"→ {hints[key]}", foreground="gray").pack(side=tk.LEFT)
        
        # Сохранение при изменении
        combobox.bind('<<ComboboxSelected>>', lambda e, k=key: self.update_combobox_value(k))
        combobox.bind('<FocusOut>', lambda e, k=key: self.update_combobox_value(k))
    
    def create_footer(self, parent):
        """Создание нижней части с кнопками"""
        footer_frame = ttk.Frame(parent)
        footer_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(footer_frame, text="?", width=3, command=self.show_help).pack(side=tk.LEFT, padx=5)
        ttk.Button(footer_frame, text="i", width=3, command=self.show_info).pack(side=tk.LEFT, padx=5)
        
        # Добавляем информацию о поддержке в нижний колонтитул
        support_label = ttk.Label(footer_frame, text=f"{self.developer_info} | Версия: {VERSION}", foreground="gray", font=('Arial', 8))
        support_label.pack(side=tk.RIGHT, padx=5)
    
    def show_column_management_dialog(self):
        """Диалог управления колонками"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Управление колонками отчета")
        dialog.geometry("400x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Управление колонками отчета:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Фрейм для списка колонок с возможностью перетаскивания
        columns_frame = ttk.LabelFrame(main_frame, text="Видимые колонки (перетащите для изменения порядка)")
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Создаем список колонок с чекбоксами
        self.column_vars = {}
        self.column_listbox = tk.Listbox(columns_frame, selectmode=tk.SINGLE)
        scrollbar = ttk.Scrollbar(columns_frame, orient=tk.VERTICAL, command=self.column_listbox.yview)
        self.column_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Заполняем список колонок
        column_names = {
            "number": "№",
            "create_time": "Время создания",
            "route": "Маршрут",
            "original_name": "Исходное имя файла",
            "new_name": "Новое имя файла"
        }
        
        for column_id in self.column_order:
            if column_id in self.column_visibility:
                display_name = column_names.get(column_id, column_id)
                status = "✓" if self.column_visibility[column_id] else "✗"
                self.column_listbox.insert(tk.END, f"{status} {display_name}")
                self.column_vars[column_id] = self.column_visibility[column_id]
        
        # Привязываем обработчик двойного клика для переключения видимости
        self.column_listbox.bind("<Double-Button-1>", self.toggle_column_visibility)
        
        # Привязываем обработчик перетаскивания для изменения порядка
        self.column_listbox.bind('<ButtonPress-1>', self.on_drag_start)
        self.column_listbox.bind('<B1-Motion>', self.on_drag_motion)
        self.column_listbox.bind('<ButtonRelease-1>', self.on_drag_release)
        
        self.column_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки управления
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="Сохранить", command=lambda: self.save_column_settings(dialog)).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Сбросить", command=self.reset_column_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Отмена", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def toggle_column_visibility(self, event):
        """Переключение видимости колонки при двойном клике"""
        selection = self.column_listbox.curselection()
        if selection:
            index = selection[0]
            column_id = self.column_order[index]
            self.column_vars[column_id] = not self.column_vars[column_id]
            
            # Обновляем отображение в списке
            column_names = {
                "number": "№",
                "create_time": "Время создания",
                "route": "Маршрут",
                "original_name": "Исходное имя файла",
                "new_name": "Новое имя файла"
            }
            
            display_name = column_names.get(column_id, column_id)
            status = "✓" if self.column_vars[column_id] else "✗"
            self.column_listbox.delete(index)
            self.column_listbox.insert(index, f"{status} {display_name}")
            self.column_listbox.selection_set(index)
    
    def on_drag_start(self, event):
        """Начало перетаскивания элемента списка"""
        self.drag_start_index = self.column_listbox.nearest(event.y)
    
    def on_drag_motion(self, event):
        """Перетаскивание элемента списка"""
        pass  # Визуальная обработка не требуется
    
    def on_drag_release(self, event):
        """Завершение перетаскивания элемента списка"""
        end_index = self.column_listbox.nearest(event.y)
        if hasattr(self, 'drag_start_index') and self.drag_start_index != end_index:
            # Перемещаем элемент в списке
            items = list(self.column_vars.items())
            item_to_move = items.pop(self.drag_start_index)
            items.insert(end_index, item_to_move)
            
            # Обновляем порядок колонок
            self.column_order = [item[0] for item in items]
            self.column_vars = dict(items)
            
            # Обновляем отображение списка
            self.column_listbox.delete(0, tk.END)
            
            column_names = {
                "number": "№",
                "create_time": "Время создания",
                "route": "Маршрут",
                "original_name": "Исходное имя файла",
                "new_name": "Новое имя файла"
            }
            
            for column_id, is_visible in items:
                display_name = column_names.get(column_id, column_id)
                status = "✓" if is_visible else "✗"
                self.column_listbox.insert(tk.END, f"{status} {display_name}")
            
            # Выделяем перемещенный элемент
            self.column_listbox.selection_set(end_index)
    
    def save_column_settings(self, dialog):
        """Сохранение настроек колонок"""
        # Сохраняем настройки видимости
        self.column_visibility = self.column_vars.copy()
        self.settings.update_setting("column_visibility", self.column_visibility)
        self.settings.update_setting("column_order", self.column_order)
        
        # Применяем изменения
        self.apply_column_visibility()
        dialog.destroy()
        logging.info("Настройки колонок сохранены")
    
    def reset_column_settings(self):
        """Сброс настроек колонок к значениям по умолчанию"""
        # Сбрасываем настройки к значениям по умолчанию
        self.column_visibility = {
            "number": True,
            "create_time": True,
            "route": True,
            "original_name": True,
            "new_name": True
        }
        self.column_order = ["number", "create_time", "route", "original_name", "new_name"]
        
        # Обновляем диалог
        self.column_listbox.delete(0, tk.END)
        
        column_names = {
            "number": "№",
            "create_time": "Время создания",
            "route": "Маршрут",
            "original_name": "Исходное имя файла",
            "new_name": "Новое имя файла"
        }
        
        for column_id in self.column_order:
            display_name = column_names.get(column_id, column_id)
            status = "✓" if self.column_visibility[column_id] else "✗"
            self.column_listbox.insert(tk.END, f"{status} {display_name}")
        
        logging.info("Настройки колонок сброшены к значениям по умолчанию")
    
    def install_plugin_dialog(self):
        """Диалог установки нового плагина"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл плагина",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
            defaultextension=".py"
        )
        
        if file_path:
            try:
                # Создаем папку plugins если нет
                plugins_dir = "plugins"
                if not os.path.exists(plugins_dir):
                    os.makedirs(plugins_dir)
                
                # Получаем имя файла
                filename = os.path.basename(file_path)
                destination = os.path.join(plugins_dir, filename)
                
                # Проверяем, не существует ли уже плагин
                if os.path.exists(destination):
                    answer = messagebox.askyesno(
                        "Подтверждение", 
                        f"Плагин '{filename}' уже существует. Заменить?"
                    )
                    if not answer:
                        return
                
                # Копируем файл
                shutil.copy2(file_path, destination)
                
                # Добавляем плагин в настройки
                plugin_name = filename[:-3]  # Убираем .py
                enabled_plugins = self.settings.settings.get("enabled_plugins", [])
                
                if plugin_name not in enabled_plugins:
                    enabled_plugins.append(plugin_name)
                    self.settings.update_setting("enabled_plugins", enabled_plugins)
                    logging.info(f"Плагин добавлен в настройки: {plugin_name}")
                
                messagebox.showinfo(
                    "Успех", 
                    f"Плагин '{plugin_name}' успешно установлен!\n\n"
                    f"Дальнейшие действия:\n"
                    f"1. Перезапустите программу\n"
                    f"2. Новая вкладка появится автоматически\n"
                    f"3. Если вкладки нет - проверьте 'Управление плагинами'"
                )
                
                logging.info(f"Плагин установлен: {filename}")
                
            except Exception as e:
                error_msg = f"Ошибка установки плагина: {e}"
                logging.error(error_msg)
                messagebox.showerror("Ошибка", error_msg)
    
    def show_plugins_dialog(self):
        """Диалог управления плагинами"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Управление плагинами")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Доступные плагины:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Фрейм для списка плагинов
        plugins_frame = ttk.Frame(main_frame)
        plugins_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Получаем список всех файлов плагинов
        plugins_dir = "plugins"
        plugin_files = []
        if os.path.exists(plugins_dir):
            plugin_files = [f for f in os.listdir(plugins_dir) 
                           if f.endswith('.py') and f != '__init__.py']
        
        enabled_plugins = self.settings.settings.get("enabled_plugins", [])
        
        plugin_vars = {}
        for plugin_file in plugin_files:
            plugin_name = plugin_file[:-3]  # Убираем .py
            var = tk.BooleanVar(value=plugin_name in enabled_plugins)
            plugin_vars[plugin_name] = var
            
            cb = ttk.Checkbutton(plugins_frame, text=plugin_name, variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        def save_plugins():
            new_enabled = [name for name, var in plugin_vars.items() if var.get()]
            self.settings.update_setting("enabled_plugins", new_enabled)
            messagebox.showinfo("Успех", "Настройки плагинов сохранены!\nПерезапустите программу для применения изменений.")
            logging.info("Настройки плагинов сохранены")
            dialog.destroy()
        
        ttk.Button(main_frame, text="Сохранить", command=save_plugins).pack(pady=10)
    
    def on_folder_selected(self):
        """Обработка выбора папки из истории"""
        if "folder_var" in self.widgets:
            folder = self.widgets["folder_var"].get()
            if folder and folder != self.settings.settings["folder"]:
                # Останавливаем мониторинг перед изменением папки
                if self.monitor:
                    self.stop_monitoring()
                
                # Обновляем настройки
                self.settings.update_setting("folder", folder)
                self.settings.add_to_folder_history(folder)
                
                # Перезапускаем мониторинг если он был активен
                if self.settings.settings.get("monitoring_enabled", False):
                    self.start_monitoring()
                
                logging.info(f"Папка изменена на: {folder}")
    
    def on_template_selected(self):
        """Обработка выбора шаблона из истории"""
        if "template_var" in self.widgets:
            template = self.widgets["template_var"].get()
            if template and template != self.settings.settings["template"]:
                self.settings.update_setting("template", template)
                self.settings.add_to_template_history(template)
                logging.info(f"Шаблон изменен на: {template}")
    
    def load_settings_to_ui(self):
        """Загрузка настроек в UI"""
        # Этот метод теперь вызывается в create_settings_tab через создание виджетов
        pass
    
    def browse_folder(self):
        """Выбор папки"""
        folder = filedialog.askdirectory()
        if folder:
            # Останавливаем мониторинг перед изменением папки
            if self.monitor:
                self.stop_monitoring()
            
            # Обновляем папку
            if "folder_var" in self.widgets:
                self.widgets["folder_var"].set(folder)
            self.settings.update_setting("folder", folder)
            self.settings.add_to_folder_history(folder)
            
            # Перезапускаем мониторинг если он был активен
            if self.settings.settings.get("monitoring_enabled", False):
                self.start_monitoring()
            
            logging.info(f"Папка изменена на: {folder}")
    
    def check_template(self):
        """Проверка шаблона"""
        try:
            example = self.generate_filename("example.jpg")
            messagebox.showinfo("Проверка шаблона", f"Пример имени файла:\n{example}")
            logging.info(f"Проверка шаблона: {example}")
        except Exception as e:
            error_msg = f"Ошибка в шаблоне: {e}"
            logging.error(error_msg)
            messagebox.showerror("Ошибка", error_msg)
    
    def update_combobox_value(self, key):
        """Обновление значения Combobox"""
        if f"{key}_var" in self.widgets:
            value = self.widgets[f"{key}_var"].get()
            self.settings.update_setting(key, value)
            
            # ДОБАВЛЕНО: добавляем новое значение в список значений комбобокса если его там нет
            self.settings.add_to_combobox_values(key, value)
            
            # ДОБАВЛЕНО: если это маршрут, обновляем фильтр в отчете
            if key == "route":
                self.update_route_filter_combobox()
            
            logging.info(f"Настройка '{key}' изменена на: {value}")
    
    def save_settings(self):
        """Сохранение настроек"""
        try:
            # Сохраняем значения из полей ввода
            if "folder_var" in self.widgets:
                folder = self.widgets["folder_var"].get()
                self.settings.update_setting("folder", folder)
                self.settings.add_to_folder_history(folder)
            
            if "ext_var" in self.widgets:
                extensions = self.widgets["ext_var"].get()
                self.settings.update_setting("extensions", extensions)
            
            if "template_var" in self.widgets:
                template = self.widgets["template_var"].get()
                self.settings.update_setting("template", template)
                self.settings.add_to_template_history(template)
            
            # Сохраняем настройку переименовывать только сегодняшние файлы
            if "rename_only_today_var" in self.widgets:
                rename_only_today = self.widgets["rename_only_today_var"].get()
                self.settings.update_setting("rename_only_today", rename_only_today)
            
            messagebox.showinfo("Успех", "Настройки сохранены!")
            logging.info("Настройки программы сохранены")
        except Exception as e:
            error_msg = f"Ошибка сохранения: {e}"
            logging.error(error_msg)
            messagebox.showerror("Ошибка", error_msg)
    
    def update_monitoring_button(self):
        """Обновление внешнего вида кнопки мониторинга и вкладки"""
        if self.monitor and self.monitor.is_monitoring:
            self.monitoring_button.config(text="ВКЛ", style="Green.TButton")
            # Изменяем текст вкладки "ЭГОК" с указанием статуса
            self.notebook.tab(self.egok_tab, text="ЭГОК 🟢 ВКЛЮЧЕН")
        else:
            self.monitoring_button.config(text="ВЫКЛ", style="Red.TButton")
            # Изменяем текст вкладки "ЭГОК" с указанием статуса
            self.notebook.tab(self.egok_tab, text="ЭГОК 🔴 ВЫКЛЮЧЕН")
    
    def toggle_monitoring(self):
        """Переключение мониторинга"""
        if self.monitor and self.monitor.is_monitoring:
            self.stop_monitoring()
            logging.info("Мониторинг отключен пользователем")
        else:
            self.start_monitoring()
            logging.info("Мониторинг включен пользователем")
        self.update_monitoring_button()

    def start_monitoring(self):
        """Запуск мониторинга - ТЕПЕРЬ БЕЗ ПЕРЕИМЕНОВАНИЯ СУЩЕСТВУЮЩИХ ФАЙЛОВ"""
        if not self.monitor:
            self.monitor = FileMonitor(self.settings, self.rename_files)
        
        success = self.monitor.start_monitoring()
        
        if success:
            # Сохраняем настройку
            self.settings.update_setting("monitoring_enabled", True)
            
            # УБРАН ВЫЗОВ ПЕРЕИМЕНОВАНИЯ СУЩЕСТВУЮЩИХ ФАЙЛОВ
            logging.info("Мониторинг запущен (без переименования существующих файлов)")
        else:
            logging.error("Не удалось запустить мониторинг")
            messagebox.showerror("Ошибка", "Не удалось запустить мониторинг")
        
        self.update_monitoring_button()

    def stop_monitoring(self):
        """Остановка мониторинга"""
        if self.monitor:
            self.monitor.stop_monitoring()
            
            # Сохраняем настройку
            self.settings.update_setting("monitoring_enabled", False)
            logging.info("Мониторинг остановлен")
        
        self.update_monitoring_button()
    
    def format_date_by_format(self, date_obj, date_format):
        """Форматирует дату по выбранному формату"""
        format_mapping = {
            "ДДММГГГГ": date_obj.strftime("%d%m%Y"),
            "ДДММГГ": date_obj.strftime("%d%m%y"),
            "ГГГГММДД": date_obj.strftime("%Y%m%d"),
            "ДД.ММ.ГГГГ": date_obj.strftime("%d.%m.%Y"),
            "ДД.ММ.ГГ": date_obj.strftime("%d.%m.%y"),
            "ГГГГ.ММ.ДД": date_obj.strftime("%Y.%m.%d")
        }
        return format_mapping.get(date_format, date_obj.strftime("%Y%m%d"))
    
    def generate_filename(self, filepath, counter=None):
        """Генерация имени файла по шаблону"""
        file_ext = Path(filepath).suffix.lower()[1:]  # Без точки
        today = datetime.now()
        
        # Если счетчик не передан, вычисляем его
        if counter is None:
            counter = self.get_next_counter()
        
        # Форматируем номер
        if self.settings.settings["number_format"] == "01":
            counter_str = f"{counter:02d}"
        elif self.settings.settings["number_format"] == "001":
            counter_str = f"{counter:03d}"
        else:
            counter_str = str(counter)
        
        # Форматируем дату по выбранному формату
        date_format = self.settings.settings.get("date_format", "ГГГГММДД")
        formatted_date = self.format_date_by_format(today, date_format)
        
        # Заменяем переменные в шаблоне
        filename = self.settings.settings["template"]
        filename = filename.replace("{project}", self.settings.settings["project"])
        filename = filename.replace("{CN}", self.settings.settings["cn_type"])
        filename = filename.replace("{route}", self.settings.settings["route"])
        filename = filename.replace("{date}", formatted_date)
        filename = filename.replace("{counter}", counter_str)
        filename = filename.replace("{extension}", file_ext)
        filename = filename.replace("{1}", self.settings.settings["var1"])
        filename = filename.replace("{2}", self.settings.settings["var2"])
        filename = filename.replace("{3}", self.settings.settings["var3"])
        
        return f"{filename}.{file_ext}"

    def get_next_counter(self):
        """Получение следующего номера счетчика с учетом уже переименованных файлов и файлов в папке"""
        folder = self.settings.settings["folder"]
        today = datetime.now()
        date_format = self.settings.settings.get("date_format", "ГГГГММДД")
        formatted_date = self.format_date_by_format(today, date_format)
        
        if not os.path.exists(folder):
            return 1
        
        # Создаем регулярное выражение для поиска файлов с текущей датой
        # Экранируем специальные символы в отформатированной дате
        escaped_date = re.escape(formatted_date)
        
        # Создаем шаблон для поиска файлов с текущей датой и форматом
        pattern = re.compile(
            f"{re.escape(self.settings.settings['project'])}_{escaped_date}_"
            f"{re.escape(self.settings.settings['route'])}_(\\d+)_{re.escape(self.settings.settings['cn_type'])}"
        )
        
        max_counter = 0
        
        for filename in os.listdir(folder):
            match = pattern.match(filename)
            if match:
                counter = int(match.group(1))
                max_counter = max(max_counter, counter)
        
        # Также проверяем историю переименований на случай, если файлы были удалены
        # но мы хотим продолжить нумерацию с правильного номера
        history_counter = self.get_max_counter_from_history()
        
        return max(max_counter, history_counter) + 1

    def get_max_counter_from_history(self):
        """Получение максимального номера из истории переименований"""
        max_counter = 0
        today = datetime.now()
        date_format = self.settings.settings.get("date_format", "ГГГГММДД")
        formatted_date = self.format_date_by_format(today, date_format)
        
        # Экранируем специальные символы в отформатированной дате
        escaped_date = re.escape(formatted_date)
        
        # Создаем шаблон для поиска в именах файлов
        pattern = re.compile(
            f"{re.escape(self.settings.settings['project'])}_{escaped_date}_"
            f"{re.escape(self.settings.settings['route'])}_(\\d+)_{re.escape(self.settings.settings['cn_type'])}"
        )
        
        for record in self.report_data:
            if len(record) > 4:  # Проверяем, что есть новое имя файла
                new_name = record[4]
                match = pattern.match(new_name)
                if match:
                    counter = int(match.group(1))
                    max_counter = max(max_counter, counter)
        
        return max_counter

    def rename_files(self, filepaths):
        """Переименование файлов с исправленной обработкой GUI"""
        try:
            with self.rename_lock:
                for filepath in filepaths:
                    try:
                        if not os.path.exists(filepath):
                            logging.warning(f"Файл не существует: {filepath}")
                            continue
                        
                        # Проверяем расширение файла
                        file_ext = Path(filepath).suffix.lower().lstrip('.')
                        extensions = [ext.strip().lower() for ext in self.settings.settings["extensions"].split(",")]
                        
                        if file_ext not in extensions:
                            logging.info(f"Файл {filepath} пропущен - расширение {file_ext} не в списке разрешенных")
                            continue
                        
                        # Проверяем, не был ли файл уже переименован программой
                        if self.renamed_files_manager.is_file_renamed(filepath):
                            logging.info(f"Файл {filepath} уже был переименован программой - пропускаем")
                            continue
                        
                        # Проверяем, нужно ли переименовывать только сегодняшние файлы
                        if self.settings.settings.get("rename_only_today", True):
                            try:
                                file_time = datetime.fromtimestamp(os.path.getctime(filepath))
                                if file_time.date() != datetime.now().date():
                                    logging.info(f"Файл {filepath} создан не сегодня - пропускаем")
                                    continue
                            except Exception as e:
                                logging.error(f"Ошибка проверки времени файла {filepath}: {e}")
                        
                        # Получаем время создания файла
                        try:
                            create_time = datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%H:%M:%S')
                        except Exception as e:
                            logging.error(f"Ошибка получения времени создания файла {filepath}: {e}")
                            create_time = datetime.now().strftime('%H:%M:%S')
                        
                        # Генерируем новое имя
                        new_name = self.generate_filename(filepath)
                        new_path = os.path.join(os.path.dirname(filepath), new_name)
                        
                        # Проверяем, не существует ли уже файл с таким именем
                        if os.path.exists(new_path):
                            logging.warning(f"Файл с именем {new_name} уже существует - пропускаем переименование")
                            continue
                        
                        # Переименовываем файл
                        os.rename(filepath, new_path)
                        
                        # Добавляем в историю переименований
                        self.renamed_files_manager.add_renamed_file(filepath)
                        
                        # ИСПРАВЛЕНИЕ БАГА: безопасный вызов добавления в отчет
                        self.root.after(0, lambda: self.add_to_report(
                            os.path.basename(filepath), new_name, new_path, create_time
                        ))
                        
                        logging.info(f"Успешно переименован: {os.path.basename(filepath)} -> {new_name}")
                        
                    except Exception as e:
                        logging.error(f"Ошибка переименования файла {filepath}: {e}")
        except Exception as e:
            logging.error(f"Общая ошибка в rename_files: {e}")
    
    def process_log_queue(self):
        """Обработка очереди логов"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                
                # Ограничиваем количество строк в логах
                self.log_line_counter += 1
                if self.log_line_counter > 1000:
                    self.log_text.delete("1.0", "100.0")
                    self.log_line_counter = 900
                
                # Добавляем сообщение с цветом в зависимости от уровня
                self.log_text.config(state=tk.NORMAL)
                
                if "ОШИБКА" in message or "ERROR" in message:
                    self.log_text.insert(tk.END, f"{message}\n", "error")
                elif "ПРЕДУПРЕЖДЕНИЕ" in message or "WARNING" in message:
                    self.log_text.insert(tk.END, f"{message}\n", "warning")
                elif "ИНФОРМАЦИЯ" in message or "INFO" in message:
                    self.log_text.insert(tk.END, f"{message}\n", "info")
                else:
                    self.log_text.insert(tk.END, f"{message}\n", "black")
                
                self.log_text.config(state=tk.DISABLED)
                self.log_text.see(tk.END)
                
        except queue.Empty:
            pass
        
        # Повторяем каждые 100 мс
        self.root.after(100, self.process_log_queue)
    
    def show_help(self):
        """Показать справку"""
        help_text = """
СПРАВКА ПО ПРОГРАММЕ EGOK RENAMER

ОСНОВНЫЕ ФУНКЦИИ:
• Автоматическое переименование файлов в указанной папке
• Мониторинг папки в реальном времени
• Гибкая настройка шаблонов имен
• Поддержка переменных в шаблонах
• Отчет о переименованных файлах

ПЕРЕМЕННЫЕ В ШАБЛОНАХ:
• {project} - название проекта
• {CN} - тип ЦН
• {route} - маршрут
• {date} - дата в выбранном формате
• {counter} - порядковый номер
• {extension} - расширение файла
• {1}, {2}, {3} - пользовательские переменные

ФОРМАТЫ ДАТЫ:
• ДДММГГГГ - 31122023
• ДДММГГ - 311223
• ГГГГММДД - 20231231
• ДД.ММ.ГГГГ - 31.12.2023
• ДД.ММ.ГГ - 31.12.23
• ГГГГ.ММ.ДД - 2023.12.31

УПРАВЛЕНИЕ ОТЧЕТОМ:
• Копирование выделенных ячеек
• Экспорт в файл
• Фильтрация по маршруту и дате
• Настройка видимости колонок

РАСШИРЕНИЯ:
• Поддержка плагинов для добавления функций
• Установка новых плагинов через диалог

ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ:
• Конструктор шаблонов для визуального создания
• История шаблонов и папок
• Настройка формата нумерации
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Справка")
        help_window.geometry("600x500")
        
        text_widget = tk.Text(help_window, wrap=tk.WORD, padx=10, pady=10)
        text_widget.insert("1.0", help_text)
        text_widget.config(state=tk.DISABLED)
        
        scrollbar = ttk.Scrollbar(help_window, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def show_info(self):
        """Показать информацию о программе"""
        info_text = f"""
EGOK RENAMER v{VERSION}

Программа для автоматического переименования файлов
по заданному шаблону с поддержкой мониторинга папки.

ФУНКЦИОНАЛ:
• Автоматическое переименование файлов
• Мониторинг папки в реальном времени
• Гибкие шаблоны имен с переменными
• Подробный отчет о действиях
• Поддержка плагинов
• Экспорт данных

РАЗРАБОТЧИК: {self.developer_info}

ИСПОЛЬЗУЕМЫЕ БИБЛИОТЕКИ:
• tkinter - графический интерфейс
• watchdog - мониторинг файловой системы
• tksheet - расширенная таблица (если доступна)
• PIL - работа с изображениями

СТАТУС МОНИТОРИНГА: {"ВКЛЮЧЕН" if (self.monitor and self.monitor.is_monitoring) else "ВЫКЛЮЧЕН"}

КОЛИЧЕСТВО ПЛАГИНОВ: {len(self.plugin_manager.plugins)}

ЗАПИСЕЙ В ОТЧЕТЕ: {len(self.report_data)}
        """
        
        messagebox.showinfo("О программе", info_text)

def main():
    """Главная функция"""
    try:
        root = tk.Tk()
        app = RenamerApp(root)
        root.mainloop()
    except Exception as e:
        logging.critical(f"Критическая ошибка: {e}")
        messagebox.showerror("Ошибка", f"Критическая ошибка при запуске: {e}")

if __name__ == "__main__":
    main()