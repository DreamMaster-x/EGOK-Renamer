# main.py
import os
import json
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
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

# Версия программы
VERSION = "3.9.0"

# Проверяем наличие tksheet
try:
    import tksheet
    TKSHEET_AVAILABLE = True
except ImportError:
    TKSHEET_AVAILABLE = False
    logging.error("Библиотека tksheet не установлена. Отчет будет ограничен в функциях.")

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
            "tl_type": "VK",
            "route": "M2.1",
            "number_format": "01",
            "var1": "Значение1",
            "var2": "Значение2",
            "var3": "Значение3",
            "folder": r"C:\video\violations",
            "extensions": "png,jpg,jpeg",
            "template": "{project}_{date}_{route}_{counter}_{TL}",
            "monitoring_enabled": True,
            "rename_only_today": True,
            "folder_history": [
                r"C:\video\violations",
                r"C:\temp\files",
                r"D:\projects\images"
            ],
            "template_history": [
                "{project}_{date}_{route}_{counter}_{TL}",
                "{project}_{TL}_{date}_{counter}",
                "{route}_{date}_{counter}_{project}"
            ],
            "enabled_plugins": ["example_plugin"],
            "combobox_values": {
                "project": ["Проект1", "Проект2"],
                "tl_type": ["VK", "Другой"],
                "route": ["M2.1", "M2.2", "M2.3"],
                "number_format": ["1", "01", "001"],
                "var1": ["Значение1", "Значение2"],
                "var2": ["Значение1", "Значение2"],
                "var3": ["Значение1", "Значение2"]
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
        self.monitor = None
        self.log_queue = queue.Queue()
        self.widgets = {}
        self.log_line_counter = 0
        self.rename_history = []
        self.current_route_filter = "Все"
        
        # Данные для отчета
        self.report_data = []
        self.filtered_report_data = []
        
        # Заголовки колонок
        self.column_headers = ["№", "Время создания", "Маршрут", "Исходное имя файла", "Новое имя файла"]
        self.column_ids = ["number", "create_time", "route", "original_name", "new_name"]
        
        # Словарь для управления видимостью колонок
        self.column_visibility = {
            "number": True,
            "create_time": True,
            "route": True,
            "original_name": True,
            "new_name": True
        }
        
        # Порядок колонок
        self.column_order = ["number", "create_time", "route", "original_name", "new_name"]
        
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
        egok_tab = ttk.Frame(self.notebook)
        self.notebook.add(egok_tab, text="ЭГОК")
        
        # Создаем содержимое вкладки ЭГОК с новой структурой
        self.create_egok_tab(egok_tab)
        
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
        
        self.create_combobox_row(scrollable_frame, "Проект:", "project", 0)
        self.create_combobox_row(scrollable_frame, "Тип ЦН:", "tl_type", 1)
        self.create_combobox_row(scrollable_frame, "Маршрут:", "route", 2)
        self.create_combobox_row(scrollable_frame, "Формат номера:", "number_format", 3)
        self.create_combobox_row(scrollable_frame, "Переменная 1:", "var1", 4)
        self.create_combobox_row(scrollable_frame, "Переменная 2:", "var2", 5)
        self.create_combobox_row(scrollable_frame, "Переменная 3:", "var3", 6)
        
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
        
        # Шаблон имени
        template_frame = ttk.Frame(rename_frame)
        template_frame.pack(fill=tk.X, pady=2)
        ttk.Label(template_frame, text="Шаблон имени:").pack(side=tk.LEFT)
        
        template_var = tk.StringVar(value=self.settings.settings["template"])
        self.widgets["template_var"] = template_var
        
        # Combobox для выбора шаблона из истории
        template_cb = ttk.Combobox(
            template_frame, 
            textvariable=template_var, 
            values=self.settings.settings.get("template_history", []),
            width=30
        )
        template_cb.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Привязываем события для обновления истории
        template_cb.bind('<<ComboboxSelected>>', lambda e: self.on_template_selected())
        template_cb.bind('<FocusOut>', lambda e: self.on_template_selected())
        
        ttk.Button(template_frame, text="Проверить", command=self.check_template).pack(side=tk.LEFT)
        
        # Опции
        options_frame = ttk.Frame(rename_frame)
        options_frame.pack(fill=tk.X, pady=2)
        
        # Переключатель мониторинга с цветовой индикацией
        monitoring_frame = ttk.Frame(options_frame)
        monitoring_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(monitoring_frame, text="Мониторинг:").pack(side=tk.LEFT)
        
        # Создаем кастомный стиль для кнопки-переключателя
        style = ttk.Style()
        style.configure("Green.TButton", background="#4CAF50", foreground="#4CAF50")
        style.configure("Red.TButton", background="#F44336", foreground="#F44336")
        
        self.monitoring_button = ttk.Button(
            monitoring_frame, 
            text="ВКЛ", 
            style="Green.TButton",
            command=self.toggle_monitoring,
            width=8
        )
        self.monitoring_button.pack(side=tk.LEFT, padx=5)
        
        # Обновляем состояние кнопки при запуске
        self.update_monitoring_button()
        
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
        
        # Фильтр по маршруту
        filter_frame = ttk.Frame(report_controls_frame)
        filter_frame.pack(side=tk.RIGHT, padx=5)
        
        ttk.Label(filter_frame, text="Фильтр по маршруту:").pack(side=tk.LEFT, padx=2)
        
        self.route_filter_var = tk.StringVar(value="Все")
        self.route_filter_cb = ttk.Combobox(
            filter_frame, 
            textvariable=self.route_filter_var,
            values=["Все"] + self.settings.settings["combobox_values"]["route"],
            state="readonly",
            width=10
        )
        self.route_filter_cb.pack(side=tk.LEFT, padx=2)
        self.route_filter_cb.bind('<<ComboboxSelected>>', self.on_route_filter_changed)
        
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
        # Создаем таблицу с включенными индексами строк
        self.report_sheet = tksheet.Sheet(
            parent,
            show_row_index=True,  # Включаем отображение индексов строк
            show_header=True,
            show_x_scrollbar=True,
            show_y_scrollbar=True,
            headers=["№", "Время создания", "Маршрут", "Исходное имя файла", "Новое имя файла"],
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
            "edit_header"
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
            # Сбрасываем фильтр
            self.route_filter_var.set("Все")
            self.current_route_filter = "Все"
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
        
        # Добавляем в таблицу
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            self.report_sheet.set_sheet_data(self.report_data)
        elif hasattr(self, 'report_tree'):
            values = (number, create_time, route, original_name, new_name)
            item_id = self.report_tree.insert("", tk.END, values=values)
            
            # Применяем текущий фильтр
            if self.current_route_filter != "Все" and route != self.current_route_filter:
                self.report_tree.detach(item_id)
            
            # Автоматически прокручиваем к последней записи
            self.report_tree.see(item_id)
        
        # Логируем добавление в отчет
        logging.info(f"Добавлено в отчет: {original_name} -> {new_name}")
    
    def apply_column_visibility(self):
        """Применить настройки видимости колонок"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            # Для tksheet скрываем/показываем колонки
            columns_to_show = []
            columns_to_hide = []
            
            for i, column in enumerate(self.column_order):
                if self.column_visibility[column]:
                    columns_to_show.append(i)
                else:
                    columns_to_hide.append(i)
            
            # ИСПРАВЛЕНИЕ: используем правильные названия методов
            if columns_to_show:
                self.report_sheet.show_columns(columns_to_show)
            if columns_to_hide:
                self.report_sheet.hide_columns(columns_to_hide)
        elif hasattr(self, 'report_tree'):
            # Для Treeview определяем видимые колонки в правильном порядке
            visible_columns = [col for col in self.column_order if self.column_visibility[col]]
            
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
                self.report_tree.heading(column, text=column_names[column])
    
    def on_route_filter_changed(self, event=None):
        """Обработка изменения фильтра по маршруту"""
        selected_route = self.route_filter_var.get()
        self.current_route_filter = selected_route
        
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            # Для tksheet фильтруем данные
            if selected_route == "Все":
                self.report_sheet.set_sheet_data(self.report_data)
            else:
                filtered_data = [row for row in self.report_data if row[2] == selected_route]
                self.report_sheet.set_sheet_data(filtered_data)
        elif hasattr(self, 'report_tree'):
            # Для Treeview показываем/скрываем элементы в соответствии с фильтром
            all_items = self.report_tree.get_children()
            
            for item in all_items:
                values = self.report_tree.item(item, "values")
                if len(values) > 2:
                    route = values[2]  # values[2] - колонка с маршрутом
                    if selected_route == "Все" or route == selected_route:
                        # Показываем элемент
                        self.report_tree.attach(item, '', 'end')
                    else:
                        # Скрываем элемент (но не удаляем)
                        self.report_tree.detach(item)
        
        logging.info(f"Применен фильтр по маршруту: {selected_route}")
    
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
            "tl_type": "{TL}",
            "route": "{route}",
            "number_format": "{counter}",
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
        dialog.geometry("300x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Видимые колонки:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Фрейм для списка колонок
        columns_frame = ttk.Frame(main_frame)
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Создаем чекбоксы для каждой колонки
        column_vars = {}
        for column in self.column_order:
            var = tk.BooleanVar(value=self.column_visibility[column])
            column_vars[column] = var
            
            # Определяем название колонки для отображения
            column_names = {
                "number": "№",
                "create_time": "Время",
                "route": "Маршрут",
                "original_name": "Исходное имя",
                "new_name": "Новое имя"
            }
            
            cb = ttk.Checkbutton(columns_frame, text=column_names[column], variable=var)
            cb.pack(anchor=tk.W, pady=2)
        
        def save_columns():
            # Сохраняем настройки видимости
            for column, var in column_vars.items():
                self.column_visibility[column] = var.get()
            
            # Применяем изменения
            self.apply_column_visibility()
            dialog.destroy()
            logging.info("Настройки видимости колонок сохранены")
        
        def reset_columns():
            # Сбрасываем настройки к значениям по умолчанию
            for column in self.column_visibility:
                self.column_visibility[column] = True
            self.column_order = ["number", "create_time", "route", "original_name", "new_name"]
            self.apply_column_visibility()
            dialog.destroy()
            logging.info("Настройки видимости колонок сброшены к значениям по умолчанию")
        
        # Фрейм для кнопок
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="Сохранить", command=save_columns).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Сбросить", command=reset_columns).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Отмена", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
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
            
            # Добавляем новое значение в список значений если его там нет
            if value not in self.settings.settings["combobox_values"][key]:
                self.settings.settings["combobox_values"][key].append(value)
                self.settings.save_settings()
            
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
        """Обновление внешнего вида кнопки мониторинга"""
        if self.monitor and self.monitor.is_monitoring:
            self.monitoring_button.config(text="ВКЛ", style="Green.TButton")
        else:
            self.monitoring_button.config(text="ВЫКЛ", style="Red.TButton")
    
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

    def stop_monitoring(self):
        """Остановка мониторинга"""
        if self.monitor:
            self.monitor.stop_monitoring()
            
            # Сохраняем настройку
            self.settings.update_setting("monitoring_enabled", False)
            logging.info("Мониторинг остановлен")
    
    def generate_filename(self, filepath, counter=None):
        """Генерация имени файла по шаблону"""
        file_ext = Path(filepath).suffix.lower()[1:]  # Без точки
        today = datetime.now().strftime("%Y%m%d")
        
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
        
        # Заменяем переменные в шаблоне
        filename = self.settings.settings["template"]
        filename = filename.replace("{project}", self.settings.settings["project"])
        filename = filename.replace("{TL}", self.settings.settings["tl_type"])
        filename = filename.replace("{route}", self.settings.settings["route"])
        filename = filename.replace("{date}", today)
        filename = filename.replace("{counter}", counter_str)
        filename = filename.replace("{extension}", file_ext)
        filename = filename.replace("{1}", self.settings.settings["var1"])
        filename = filename.replace("{2}", self.settings.settings["var2"])
        filename = filename.replace("{3}", self.settings.settings["var3"])
        
        return f"{filename}.{file_ext}"

    def get_next_counter(self):
        """Получение следующего номера счетчика с учетом уже переименованных файлов и файлов в папке"""
        folder = self.settings.settings["folder"]
        today = datetime.now().strftime("%Y%m%d")
        
        if not os.path.exists(folder):
            return 1
        
        # Ищем все файлы с сегодняшней датой и соответствующим профилем
        pattern = re.compile(
            f"{re.escape(self.settings.settings['project'])}_{re.escape(today)}_"
            f"{re.escape(self.settings.settings['route'])}_(\\d+)_{re.escape(self.settings.settings['tl_type'])}"
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
        today = datetime.now().strftime("%Y%m%d")
        
        pattern = re.compile(
            f"{re.escape(self.settings.settings['project'])}_{re.escape(today)}_"
            f"{re.escape(self.settings.settings['route'])}_(\\d+)_{re.escape(self.settings.settings['tl_type'])}"
        )
        
        # Проверяем файлы в папке, которые уже соответствуют шаблону
        folder = self.settings.settings["folder"]
        if os.path.exists(folder):
            for filename in os.listdir(folder):
                match = pattern.match(filename)
                if match:
                    counter = int(match.group(1))
                    max_counter = max(max_counter, counter)
        
        return max_counter

    def rename_files(self, specific_files=None):
        """Переименование файлов с улучшенной логикой - ТОЛЬКО НОВЫЕ ФАЙЛЫ"""
        try:
            # Используем блокировку для предотвращения конфликтов при одновременном переименовании
            with self.rename_lock:
                folder = self.settings.settings["folder"]
                extensions = [ext.strip().lower() for ext in self.settings.settings["extensions"].split(",")]
                
                if not os.path.exists(folder):
                    logging.error(f"Папка не существует: {folder}")
                    return
                
                # Если переданы конкретные файлы (при мониторинге) - обрабатываем только их
                if specific_files:
                    files_to_process = []
                    for filepath in specific_files:
                        if os.path.exists(filepath):
                            file_ext = Path(filepath).suffix.lower()[1:]
                            if file_ext in extensions:
                                # Проверяем, нужно ли переименовывать только сегодняшние файлы
                                if self.settings.settings.get("rename_only_today", True):
                                    if self.is_file_from_today(filepath):
                                        files_to_process.append(filepath)
                                    else:
                                        logging.info(f"Файл {filepath} пропущен - создан не сегодня")
                                else:
                                    files_to_process.append(filepath)
                            else:
                                logging.info(f"Файл {filepath} пропущен - расширение {file_ext} не в списке разрешенных")
                
                # УБРАНА ОБРАБОТКА СУЩЕСТВУЮЩИХ ФАЙЛОВ ПРИ ЗАПУСКЕ МОНИТОРИНГА
                
                # Сортируем файлы по времени создания для правильной нумерации
                files_to_process.sort(key=lambda x: os.path.getctime(x))
                
                # Получаем начальный счетчик
                start_counter = self.get_next_counter()
                current_counter = start_counter
                
                for filepath in files_to_process:
                    try:
                        # Проверяем, не был ли файл уже переименован программой
                        if self.renamed_files_manager.is_file_renamed(filepath):
                            logging.info(f"Файл {filepath} уже был переименован программой, пропускаем")
                            continue
                        
                        # Получаем время создания файла ДО переименования                      
                        try:
                            create_time = datetime.fromtimestamp(os.path.getctime(filepath)).strftime('%H:%M:%S')
                        except Exception as e:
                            logging.error(f"Ошибка получения времени создания файла {filepath}: {e}")
                            create_time = datetime.now().strftime('%H:%M:%S')

                        original_name = Path(filepath).name
                        new_name = self.generate_filename(filepath, current_counter)
                        new_path = os.path.join(folder, new_name)

                        # Проверяем, не переименован ли уже файл
                        if not os.path.exists(new_path):
                            # Добавляем задержку для гарантии, что файл полностью доступен
                            time.sleep(0.1)
                            os.rename(filepath, new_path)
                            log_message = f"{original_name} -> {new_name}"
                            logging.info(log_message)
                            self.log_queue.put(log_message)
                            
                            # Добавляем в отчет с временем создания
                            self.root.after(0, lambda ct=create_time: self.add_to_report(original_name, new_name, filepath, ct))
                            
                            # Добавляем файл в историю переименований
                            self.renamed_files_manager.add_renamed_file(filepath)
                            current_counter += 1
                        else:
                            # Если файл с таким именем уже существует, ищем следующий свободный номер
                            conflict_counter = current_counter + 1
                            while os.path.exists(new_path):
                                new_name = self.generate_filename(filepath, conflict_counter)
                                new_path = os.path.join(folder, new_name)
                                conflict_counter += 1
                            
                            # Добавляем задержку для гарантии, что файл полностью доступен
                            time.sleep(0.1)
                            os.rename(filepath, new_path)
                            log_message = f"{original_name} -> {new_name} (автоподбор номера)"
                            logging.info(log_message)
                            self.log_queue.put(log_message)
                            
                            # Добавляем в отчет с временем создания
                            self.root.after(0, lambda ct=create_time: self.add_to_report(original_name, new_name, filepath, ct))
                            
                            # Добавляем файл в историю переименований
                            self.renamed_files_manager.add_renamed_file(filepath)
                            current_counter = conflict_counter
                    
                    except Exception as e:
                        error_msg = f"Ошибка переименования {filepath}: {e}"
                        logging.error(error_msg)
                        self.log_queue.put(f"ОШИБКА: {error_msg}")
        
        except Exception as e:
            error_msg = f"Общая ошибка переименования: {e}"
            logging.error(error_msg)
            self.log_queue.put(f"ОШИБКА: {error_msg}")

    def is_already_renamed(self, filename):
        """Проверяет, соответствует ли имя файла шаблону переименования"""
        today = datetime.now().strftime("%Y%m%d")
        pattern = re.compile(
            f"{re.escape(self.settings.settings['project'])}_{re.escape(today)}_"
            f"{re.escape(self.settings.settings['route'])}_(\\d+)_{re.escape(self.settings.settings['tl_type'])}"
        )
        return pattern.match(filename) is not None

    def is_file_from_today(self, filepath):
        """Проверка, создан ли файл сегодня"""
        try:
            create_time = datetime.fromtimestamp(os.path.getctime(filepath))
            return create_time.date() == datetime.now().date()
        except:
            return False

    def process_log_queue(self):
        """Обработка очереди логов с чередованием цветов"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.config(state=tk.NORMAL)
                
                # Определяем цвет для текущей строки
                if "ОШИБКА" in message:
                    color_tag = "error"
                elif "ПРЕДУПРЕЖДЕНИЕ" in message or "WARNING" in message:
                    color_tag = "warning"
                elif "ИНФОРМАЦИЯ" in message or "INFO" in message:
                    color_tag = "info"
                elif self.log_line_counter % 2 == 0:
                    color_tag = "black"
                else:
                    color_tag = "gray"
                
                # Вставляем сообщение с нужным цветом
                self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n", color_tag)
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
                
                # Увеличиваем счетчик
                self.log_line_counter += 1
                
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_log_queue)
    
    def show_help(self):
        """Показать справку"""
        help_text = f"""УЛУЧШЕННОЕ УПРАВЛЕНИЕ ОТЧЕТОМ:

ВКЛАДКА ОТЧЕТА ТЕПЕРЬ ПОДДЕРЖИВАЕТ:
✓ ВЫДЕЛЕНИЕ ОТДЕЛЬНЫХ ЯЧЕЕК - клик по ячейке
✓ ВЫДЕЛЕНИЕ ДИАПАЗОНА ЯЧЕЕК - перетаскивание мышью
✓ ВЫДЕЛЕНИЕ ЦЕЛЫХ СТРОК - клик по номеру строки слева
✓ ВЫДЕЛЕНИЕ ЦЕЛЫХ СТОЛБЦОВ - клик по заголовку столбца
✓ МНОЖЕСТВЕННОЕ ВЫДЕЛЕНИЕ - Ctrl+клик для добавления
✓ ВЫДЕЛЕНИЕ ПРЯМОУГОЛЬНЫХ ОБЛАСТЕЙ - Shift+клик

ГОРЯЧИЕ КЛАВИШИ:
Ctrl+A - Выделить всю таблицу
Ctrl+C - Копировать выделенное
Ctrl+V - Вставить данные
Delete - Очистить выделенные ячейки

КОНТЕКСТНОЕ МЕНЮ (ПРАВАЯ КНОПКА МЫШИ):
• Копировать выделенное
• Копировать как текст  
• Выделить строку/столбец
• Скрыть/показать столбцы
• Экспорт выделенного
• Очистить отчет

ФУНКЦИИ КОПИРОВАНИЯ:
- Автоматическое форматирование при копировании
- Поддержка вставки в Excel и другие табличные редакторы
- Сохранение структуры данных
- Подсчет скопированных элементов

УПРАВЛЕНИЕ СТОЛБЦАМИ:
- Настройка видимости столбцов
- Изменение порядка столбцов
- Скрытие/отображение столбцов
- Сброс к настройкам по умолчанию

ОСНОВНЫЕ ФУНКЦИИ ПРОГРАММЫ:
- Автоматическое переименование НОВЫХ файлов
- Мониторинг папки в реальном времени
- Гибкие шаблоны имен

ТЕХНИЧЕСКАЯ ПОДДЕРЖКА:
Telegram: @xDream_Master
Email: drea_m_aster@vk.com
Версия: {VERSION}"""
    
        messagebox.showinfo("Справка", help_text)
        logging.info("Открыта справка")
    
    def show_info(self):
        """Показать информацию о программе"""
        info_text = f"""EGOK Renamer v{VERSION}

{self.developer_info}

ОСНОВНЫЕ УЛУЧШЕНИЯ ВЕРСИИ {VERSION}:

🎯 УЛУЧШЕННОЕ УПРАВЛЕНИЕ ОТЧЕТОМ:
• Excel-подобное выделение ячеек, строк и столбцов
• Перетаскивание для выделения диапазонов
• Множественное выделение (Ctrl+клик)
• Выделение прямоугольных областей (Shift+клик)
• Контекстное меню с расширенными функциями

📋 РАСШИРЕННЫЕ ВОЗМОЖНОСТИ КОПИРОВАНИЯ:
• Копирование отдельных ячеек
• Копирование диапазонов
• Копирование целых строк/столбцов
• Форматированное копирование для Excel
• Экспорт выделенных данных в CSV

⚙️ УПРАВЛЕНИЕ СТОЛБЦАМИ:
• Скрытие/отображение столбцов
• Настройка видимости через диалог
• Сброс настроек к значениям по умолчанию

Техническая поддержка:
- Telegram: @xDream_Master
- Email: drea_m_aster@vk.com

© 2024 Все права защищены."""
    
        messagebox.showinfo("О программе", info_text)
        logging.info("Открыта информация о программе")
    
    def on_closing(self):
        """Обработка закрытия программы"""
        logging.info("=" * 50)
        logging.info("ЗАВЕРШЕНИЕ РАБОТЫ ПРОГРАММЫ")
        logging.info("=" * 50)
        
        if self.monitor:
            self.stop_monitoring()
        # Сохраняем историю переименований при закрытии
        self.renamed_files_manager.save_history()
        self.root.destroy()

def main():
    """Главная функция"""
    # Добавляем обработку непойманных исключений
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        logging.critical("Необработанное исключение:", exc_info=(exc_type, exc_value, exc_traceback))
        messagebox.showerror("Критическая ошибка", 
                           f"Произошла критическая ошибка:\n{exc_value}\n\n"
                           f"Подробности в файле лога.")
    
    sys.excepthook = handle_exception
    
    root = tk.Tk()
    app = RenamerApp(root)
    
    # Обработка закрытия окна
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.mainloop()

if __name__ == "__main__":
    main()