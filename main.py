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
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from PIL import Image, ImageTk

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
    """Менеджер плагинов для загрузки дополнительных вкладок"""
    
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
                    threading.Timer(0.5, lambda: self.rename_callback([event.src_path])).start()
                else:
                    logging.info(f"Файл {event.src_path} пропущен - расширение {file_ext} не в списке разрешенных")

class RenamerApp:
    """Главное приложение"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("EGOK Renamer")
        self.root.geometry("900x600")
        
        # Установка иконки приложения
        self.set_app_icon()
        
        # Добавляем информацию о разработчике в заголовок
        self.developer_info = "Разработчик: github.com/DreamMaster-x"
        
        self.settings = Settings()
        self.monitor = None
        self.log_queue = queue.Queue()
        self.widgets = {}
        self.log_line_counter = 0  # Счетчик для чередования цветов
        self.renamed_files = set()  # Множество для хранения переименованных файлов
        
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
        """Настройка логирования"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('renamer.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
    
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
        
        title_label = ttk.Label(title_frame, text="EGOK Renamer", font=('Arial', 16, 'bold'))
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
        
        # Нотебук для основных функций и плагинов
        self.notebook = ttk.Notebook(notebook_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Основная вкладка "ЭГОК" с настройками и логами
        egok_tab = ttk.Frame(self.notebook)
        self.notebook.add(egok_tab, text="ЭГОК")
        
        # Создаем содержимое вкладки ЭГОК с разделителем
        self.create_egok_tab(egok_tab)
        
        # Загружаем и создаем вкладки плагинов
        self.plugin_manager.load_plugins()
        self.plugin_manager.create_plugin_tabs(self.notebook)
    
    def create_egok_tab(self, parent):
        """Создание основной вкладки ЭГОК с настройками и логами"""
        # Создаем PanedWindow для разделения на левую и правую часть
        paned_window = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)
        
        # Левая часть - настройки
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=1)
        
        # Правая часть - логи
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)
        
        # Создаем содержимое левой части (настройки)
        self.create_settings_tab(left_frame)
        
        # Создаем содержимое правой части (логи)
        self.create_log_tab(right_frame)
        
        # Устанавливаем начальное соотношение размеров (60% настройки, 40% логи)
        paned_window.sashpos(0, int(parent.winfo_reqwidth() * 0.6))
    
    def create_settings_tab(self, parent):
        """Создание левой части с настройками"""
        # Основные настройки
        main_settings_frame = ttk.LabelFrame(parent, text="Основные настройки")
        main_settings_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Создаем скроллируемую область для настроек
        settings_canvas = tk.Canvas(main_settings_frame)
        scrollbar = ttk.Scrollbar(main_settings_frame, orient="vertical", command=settings_canvas.yview)
        scrollable_frame = ttk.Frame(settings_canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: settings_canvas.configure(scrollregion=settings_canvas.bbox("all"))
        )
        
        settings_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        settings_canvas.configure(yscrollcommand=scrollbar.set)
        
        settings_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
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
        style.configure("Red.TButton", background="#F44336", foreground="red")
        
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
        
        # Кнопки управления
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.pack(pady=10, fill=tk.X)
        
        ttk.Button(button_frame, text="Сохранить настройки", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Управление плагинами", command=self.show_plugins_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Установить плагин", command=self.install_plugin_dialog).pack(side=tk.LEFT, padx=5)
    
    def create_log_tab(self, parent):
        """Создание правой части с логами"""
        log_frame = ttk.LabelFrame(parent, text="Логи программы")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)
        
        # Текстовое поле для логов с поддержкой цветов
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.tag_configure("black", foreground="black")
        self.log_text.tag_configure("gray", foreground="gray")
        self.log_text.tag_configure("error", foreground="red")
        
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
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
        support_label = ttk.Label(footer_frame, text=self.developer_info, foreground="gray", font=('Arial', 8))
        support_label.pack(side=tk.RIGHT, padx=5)
    
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
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка в шаблоне: {e}")
    
    def update_combobox_value(self, key):
        """Обновление значения Combobox"""
        if f"{key}_var" in self.widgets:
            value = self.widgets[f"{key}_var"].get()
            self.settings.update_setting(key, value)
            
            # Добавляем новое значение в список значений если его там нет
            if value not in self.settings.settings["combobox_values"][key]:
                self.settings.settings["combobox_values"][key].append(value)
                self.settings.save_settings()
    
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
            
            if "monitoring_var" in self.widgets:
                monitoring_enabled = self.widgets["monitoring_var"].get()
                self.settings.update_setting("monitoring_enabled", monitoring_enabled)
            
            messagebox.showinfo("Успех", "Настройки сохранены!")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения: {e}")
    
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
        else:
            self.start_monitoring()
        self.update_monitoring_button()

    def start_monitoring(self):
        """Запуск мониторинга"""
        if not self.monitor:
            self.monitor = FileMonitor(self.settings, self.rename_files)
        
        success = self.monitor.start_monitoring()
        
        if success:
            # Сохраняем настройку
            self.settings.update_setting("monitoring_enabled", True)
            logging.info("Мониторинг запущен")
        else:
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
        """Получение следующего номера счетчика с учетом уже переименованных файлов"""
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
        
        return max_counter + 1

    def rename_files(self, specific_files=None):
        """Переименование файлов"""
        try:
            folder = self.settings.settings["folder"]
            extensions = [ext.strip().lower() for ext in self.settings.settings["extensions"].split(",")]
            
            if not os.path.exists(folder):
                logging.error(f"Папка не существует: {folder}")
                return
            
            if specific_files:
                # Если переданы конкретные файлы (при мониторинге) - фильтруем по расширениям
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
            else:
                # При автоматическом переименовании - обрабатываем все подходящие файлы
                files_to_process = []
                for filename in os.listdir(folder):
                    filepath = os.path.join(folder, filename)
                    if (os.path.isfile(filepath) and 
                        Path(filename).suffix.lower()[1:] in extensions and
                        not self.is_already_renamed(filename)):
                        
                        # Проверяем, нужно ли переименовывать только сегодняшние файлы
                        if self.settings.settings.get("rename_only_today", True):
                            if self.is_file_from_today(filepath):
                                files_to_process.append(filepath)
                            else:
                                logging.info(f"Файл {filepath} пропущен - создан не сегодня")
                        else:
                            files_to_process.append(filepath)
            
            # Сортируем файлы по времени создания для правильной нумерации
            files_to_process.sort(key=lambda x: os.path.getctime(x))
            
            # Получаем начальный счетчик
            start_counter = self.get_next_counter()
            current_counter = start_counter
            
            for filepath in files_to_process:
                try:
                    # Проверяем, не был ли файл уже переименован
                    if filepath in self.renamed_files:
                        continue
                    
                    new_name = self.generate_filename(filepath, current_counter)
                    new_path = os.path.join(folder, new_name)
                    
                    # Проверяем, не переименован ли уже файл
                    if not os.path.exists(new_path):
                        os.rename(filepath, new_path)
                        log_message = f"{Path(filepath).name} -> {new_name}"
                        logging.info(log_message)
                        self.log_queue.put(log_message)
                        
                        # Добавляем файл в множество переименованных
                        self.renamed_files.add(filepath)
                        current_counter += 1
                    else:
                        # Если файл с таким именем уже существует, ищем следующий свободный номер
                        conflict_counter = current_counter + 1
                        while os.path.exists(new_path):
                            new_name = self.generate_filename(filepath, conflict_counter)
                            new_path = os.path.join(folder, new_name)
                            conflict_counter += 1
                        
                        os.rename(filepath, new_path)
                        log_message = f"{Path(filepath).name} -> {new_name} (автоподбор номера)"
                        logging.info(log_message)
                        self.log_queue.put(log_message)
                        
                        # Добавляем файл в множество переименованных
                        self.renamed_files.add(filepath)
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
        help_text = """КАК ПОЛЬЗОВАТЬСЯ ПРОГРАММОЙ:

1. ОСНОВНЫЕ ФУНКЦИИ:
- Автоматическое переименование файлов
- Мониторинг папки в реальном времени
- Гибкие шаблоны имен

2. УСТАНОВКА ПЛАГИНОВ:
❶ Автоматическая установка:
   - Настройки → "Установить плагин"
   - Выбете файл .py
   - Перезапустите программу

❷ Ручная установка:
   - Скопируйте файл .py в папку plugins/
   - Настройки → "Управление плагинами" 
   - Отметьте плагин галочкой
   - Сохраните и перезапустите

3. УПРАВЛЕНИЕ МОНИТОРИНГОМ:
- Зеленая кнопка "ВКЛ" - мониторинг активен
- Красная кнопка "ВЫКЛ" - мониторинг остановлен

4. ШАБЛОНЫ ИМЕН:
Доступные переменные:
{project} - название проекта
{route} - маршрут
{TL} - тип ЦН  
{date} - дата (ГГГГММДД)
{counter} - порядковый номер
{extension} - расширение файла
{1}, {2}, {3} - дополнительные переменные

5. ТЕХНИЧЕСКАЯ ПОДДЕРЖКА:
https://github.com/DreamMaster-x
Telegram: @xDream_Master
Email: drea_m_aster@vk.com"""
    
        messagebox.showinfo("Справка", help_text)
    
    def show_info(self):
        """Показать информацию о программе"""
        info_text = f"""EGOK Renamer v3.1

{self.developer_info}

Техническая поддержка:
- Telegram: @xDream_Master
- Email: drea_m_aster@vk.com

НОВЫЕ ФУНКЦИИ ВЕРСИИ 3.1:
- Объединенный интерфейс "ЭГОК" с настройками и логами
- Регулируемый разделитель между настройками и логами
- Модульная архитектура с поддержкой плагинов
- История папок и шаблонов
- Цветной переключатель мониторинга

СИСТЕМА ПЛАГИНОВ:
Разработчики могут добавлять новые функции через плагины:
1. Создать файл в папке plugins/
2. Унаследоваться от BasePlugin
3. Реализовать методы get_tab_name() и create_tab()

Функции:
- Автоматическое переименование файлов
- Управление мониторингом (включение/выключение)
- Гибкие настройки шаблонов
- Фильтрация файлов по расширениям
- Подробное логирование с цветным выделением

© 2024 Все права защищены."""

        messagebox.showinfo("О программе", info_text)
    
    def on_closing(self):
        """Обработка закрытия программы"""
        if self.monitor:
            self.stop_monitoring()
        self.root.destroy()

def main():
    """Главная функция"""
    root = tk.Tk()
    app = RenamerApp(root)
    
    # Обработка закрытия окна
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.mainloop()

if __name__ == "__main__":
    main()