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
            "enabled_plugins": ["example_plugin", "report_plugin"],
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
        
        # Блокировка для безопасного доступа к общим ресурсам
        self.rename_lock = threading.Lock()
        
        # Инициализация менеджера плагинов
        self.plugin_manager = PluginManager(self.settings, self.root)
        
        self.setup_logging()
        self.create_widgets()
        self.load_settings_to_ui()
        self.process_log_queue()
        
        # Запуск мониторинга если включен
        if self.settings.settings.get("monitoring_enabled", True):
            self.start_monitoring()
        
        # Обновляем состояние кнопки и вкладки после инициализации
        self.update_monitoring_button()
        
        # Добавляем начальное сообщение в логи
        self.add_log_message("Программа успешно запущена", "info")
        self.add_log_message(f"Версия: {VERSION}", "info")
        self.add_log_message(f"Папка мониторинга: {self.settings.settings['folder']}", "info")
        self.add_log_message(f"Статус мониторинга: {'ВКЛЮЧЕН' if (self.monitor and self.monitor.is_monitoring) else 'ВЫКЛЮЧЕН'}", "info")
    
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
        """Настройка логирования в txt файл"""
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
        
        # Добавляем обработчики
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
        
        logging.info("=" * 50)
        logging.info("ПРОГРАММА ЗАПУЩЕНА")
        logging.info("=" * 50)
        logging.info(f"Логирование инициализировано: {log_filename}")
    
    def add_log_message(self, message, level="info"):
        """Добавление сообщения в лог программы"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        
        # Добавляем в очередь для отображения в GUI
        self.log_queue.put((formatted_message, level))
        
        # Также логируем стандартным способом
        if level == "error":
            logging.error(message)
        elif level == "warning":
            logging.warning(message)
        else:
            logging.info(message)
    
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
                self.add_log_message("Логотип загружен успешно", "info")
            else:
                # Заглушка если изображение не найдено
                logo_label = ttk.Label(header_frame, text="[Логотип]", width=20)
                logo_label.pack(side=tk.RIGHT)
                self.add_log_message("Файл логотипа background.png не найден", "warning")
                
        except Exception as e:
            self.add_log_message(f"Ошибка загрузки логотипа: {e}", "error")
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
        """Создание основной вкладки ЭГОК БЕЗ отчета"""
        # УПРОЩЕННАЯ ВЕРСИЯ - только настройки и логи
        
        # Главный фрейм для настроек и логов
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Создаем содержимое настроек и логов
        self.create_settings_and_logs_tab(main_frame)

    def create_settings_and_logs_tab(self, parent):
        """Создание вкладки с настройками и логами"""
        # Создаем вертикальный PanedWindow для разделения настроек и логов
        paned_window = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        # Верхняя часть - основные настройки
        settings_frame = ttk.LabelFrame(paned_window, text="Основные настройки")
        paned_window.add(settings_frame, weight=2)
        
        # Нижняя часть - логи
        log_frame = ttk.LabelFrame(paned_window, text="Логи программы")
        paned_window.add(log_frame, weight=1)
        
        # Заполняем фреймы
        self.create_settings_content(settings_frame)
        self.create_log_content(log_frame)
        
        # Устанавливаем соотношение (70% настройки, 30% логи)
        paned_window.sashpos(0, int(parent.winfo_reqheight() * 0.7))
    
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
        
        # НОВАЯ КНОПКА - КОНСТРУКТОР ШАБЛОна
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
            
            self.add_log_message(f"Шаблон обновлен через конструктор: {dialog.result_template}", "info")
            
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
        self.log_text.tag_configure("success", foreground="green")
        
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
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
                
                self.add_log_message(f"Папка изменена на: {folder}", "info")
    
    def on_template_selected(self):
        """Обработка выбора шаблона из истории"""
        if "template_var" in self.widgets:
            template = self.widgets["template_var"].get()
            if template and template != self.settings.settings["template"]:
                self.settings.update_setting("template", template)
                self.settings.add_to_template_history(template)
                self.add_log_message(f"Шаблон изменен на: {template}", "info")
    
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
            
            self.add_log_message(f"Папка изменена на: {folder}", "info")
    
    def check_template(self):
        """Проверка шаблона"""
        try:
            example = self.generate_filename("example.jpg")
            messagebox.showinfo("Проверка шаблона", f"Пример имени файла:\n{example}")
            self.add_log_message(f"Проверка шаблона: {example}", "info")
        except Exception as e:
            error_msg = f"Ошибка в шаблоне: {e}"
            self.add_log_message(error_msg, "error")
            messagebox.showerror("Ошибка", error_msg)
    
    def update_combobox_value(self, key):
        """Обновление значения Combobox"""
        if f"{key}_var" in self.widgets:
            value = self.widgets[f"{key}_var"].get()
            self.settings.update_setting(key, value)
            
            # ДОБАВЛЕНО: добавляем новое значение в список значений комбобокса если его там нет
            self.settings.add_to_combobox_values(key, value)
            
            self.add_log_message(f"Настройка '{key}' изменена на: {value}", "info")
    
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
            self.add_log_message("Настройки программы сохранены", "info")
        except Exception as e:
            error_msg = f"Ошибка сохранения: {e}"
            self.add_log_message(error_msg, "error")
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
            self.add_log_message("Мониторинг отключен пользователем", "info")
        else:
            self.start_monitoring()
            self.add_log_message("Мониторинг включен пользователем", "info")
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
            self.add_log_message("Мониторинг запущен (без переименования существующих файлов)", "info")
        else:
            self.add_log_message("Не удалось запустить мониторинг", "error")
            messagebox.showerror("Ошибка", "Не удалось запустить мониторинг")
        
        self.update_monitoring_button()

    def stop_monitoring(self):
        """Остановка мониторинга"""
        if self.monitor:
            self.monitor.stop_monitoring()
            
            # Сохраняем настройку
            self.settings.update_setting("monitoring_enabled", False)
            self.add_log_message("Мониторинг остановлен", "info")
        
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
        
        # Ищем в базе данных
        records = self.db_manager.get_records_by_date()
        for record in records:
            new_name = record['new_name']
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
                            self.add_log_message(f"Файл не существует: {filepath}", "warning")
                            continue
                        
                        # Проверяем расширение файла
                        file_ext = Path(filepath).suffix.lower().lstrip('.')
                        extensions = [ext.strip().lower() for ext in self.settings.settings["extensions"].split(",")]
                        
                        if file_ext not in extensions:
                            self.add_log_message(f"Файл {filepath} пропущен - расширение {file_ext} не в списке разрешенных", "info")
                            continue
                        
                        # Проверяем, не был ли файл уже переименован программой
                        if self.renamed_files_manager.is_file_renamed(filepath):
                            self.add_log_message(f"Файл {filepath} уже был переименован программой - пропускаем", "info")
                            continue
                        
                        # Проверяем, нужно ли переименовывать только сегодняшние файлы
                        if self.settings.settings.get("rename_only_today", True):
                            try:
                                file_time = datetime.fromtimestamp(os.path.getctime(filepath))
                                if file_time.date() != datetime.now().date():
                                    self.add_log_message(f"Файл {filepath} создан не сегодня - пропускаем", "info")
                                    continue
                            except Exception as e:
                                self.add_log_message(f"Ошибка проверки времени файла {filepath}: {e}", "error")
                        
                        # Получаем время создания файла
                        try:
                            create_time = datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%H:%M:%S')
                        except Exception as e:
                            self.add_log_message(f"Ошибка получения времени создания файла {filepath}: {e}", "error")
                            create_time = datetime.now().strftime('%H:%M:%S')
                        
                        # Генерируем новое имя
                        new_name = self.generate_filename(filepath)
                        new_path = os.path.join(os.path.dirname(filepath), new_name)
                        
                        # Проверяем, не существует ли уже файл с таким именем
                        if os.path.exists(new_path):
                            self.add_log_message(f"Файл с именем {new_name} уже существует - пропускаем переименование", "warning")
                            continue
                        
                        # Переименовываем файл
                        os.rename(filepath, new_path)
                        
                        # Добавляем в историю переименований
                        self.renamed_files_manager.add_renamed_file(filepath)
                        
                        # Сохраняем в базу данных
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        route = self.settings.settings["route"]
                        self.db_manager.add_record(timestamp, route, os.path.basename(filepath), new_name, new_path)
                        
                        # Передаем данные в плагин отчетов (если он загружен)
                        self.add_to_report(os.path.basename(filepath), new_name, new_path, create_time)
                        
                        # ДОБАВЛЕНО: логирование успешного переименования
                        self.add_log_message(f"Успешно переименован: {os.path.basename(filepath)} -> {new_name}", "success")
                        
                    except Exception as e:
                        self.add_log_message(f"Ошибка переименования файла {filepath}: {e}", "error")
        except Exception as e:
            self.add_log_message(f"Общая ошибка в rename_files: {e}", "error")
    
    def add_to_report(self, original_name, new_name, filepath, create_time=None):
        """Добавление записи в отчет через плагин (если он загружен)"""
        # Ищем плагин отчетов среди загруженных плагинов
        report_plugin = None
        for plugin_name, plugin_instance in self.plugin_manager.plugins.items():
            if hasattr(plugin_instance, 'add_to_report'):
                report_plugin = plugin_instance
                break
        
        # Если нашли плагин отчетов, передаем данные ему
        if report_plugin:
            report_plugin.add_to_report(original_name, new_name, filepath, create_time)
        else:
            # Логируем, если плагин не найден
            self.add_log_message("Плагин отчетов не загружен, запись не добавлена в отчет", "warning")
    
    def process_log_queue(self):
        """Обработка очереди логов"""
        try:
            while True:
                message, level = self.log_queue.get_nowait()
                
                # Ограничиваем количество строк в логах
                self.log_line_counter += 1
                if self.log_line_counter > 1000:
                    self.log_text.delete("1.0", "100.0")
                    self.log_line_counter = 900
                
                # Добавляем сообщение с цветом в зависимости от уровня
                self.log_text.config(state=tk.NORMAL)
                
                if level == "error":
                    self.log_text.insert(tk.END, f"{message}\n", "error")
                elif level == "warning":
                    self.log_text.insert(tk.END, f"{message}\n", "warning")
                elif level == "info":
                    self.log_text.insert(tk.END, f"{message}\n", "info")
                elif level == "success":
                    self.log_text.insert(tk.END, f"{message}\n", "success")
                else:
                    self.log_text.insert(tk.END, f"{message}\n", "black")
                
                self.log_text.config(state=tk.DISABLED)
                self.log_text.see(tk.END)
                
        except queue.Empty:
            pass
        
        # Повторяем каждые 100 мс
        self.root.after(100, self.process_log_queue)
    
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
                    self.add_log_message(f"Плагин добавлен в настройки: {plugin_name}", "info")
                
                messagebox.showinfo(
                    "Успех", 
                    f"Плагин '{plugin_name}' успешно установлен!\n\n"
                    f"Дальнейшие действия:\n"
                    f"1. Перезапустите программу\n"
                    f"2. Новая вкладка появится автоматически\n"
                    f"3. Если вкладки нет - проверьте 'Управление плагинами'"
                )
                
                self.add_log_message(f"Плагин установлен: {filename}", "info")
                
            except Exception as e:
                error_msg = f"Ошибка установки плагина: {e}"
                self.add_log_message(error_msg, "error")
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
            self.add_log_message("Настройки плагинов сохранены", "info")
            dialog.destroy()
        
        ttk.Button(main_frame, text="Сохранить", command=save_plugins).pack(pady=10)
    
    def show_help(self):
        """Показать справку"""
        help_text = """
СПРАВКА ПО ПРОГРАММЕ EGOK RENAMER

ОСНОВНЫЕ ФУНКЦИИ:
• Автоматическое переименование файлов в указанной папке
• Мониторинг папки в реальном времени
• Гибкая настройка шаблонов имен
• Поддержка переменных в шаблонах
• Отчет о переименованных файлах (в плагине)

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
• Отчеты вынесены в отдельный плагин
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
• Подробный отчет о действиях (в плагине)
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