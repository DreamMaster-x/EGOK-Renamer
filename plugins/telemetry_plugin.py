# plugins/telemetry_plugin.py
import os
import json
import zipfile
import threading
import queue
import logging
import serial
import serial.tools.list_ports
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS
import math

class TelemetryPlugin:
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.log_queue = queue.Queue()
        self.serial_connection = None
        self.is_reading_telemetry = False
        self.setup_plugin_settings()
    
    def setup_plugin_settings(self):
        """Инициализация настроек плагина"""
        plugin_settings = self.settings.settings.get("telemetry_plugin", {})
        
        # Настройки по умолчанию для плагина
        default_plugin_settings = {
            "telemetry_folder": "",
            "photos_folder": "",
            "relief_folder": "",
            "output_telemetry_name": "tele_photo.tlm",
            "archive_template": "Дата_{номер маршрута}",
            "compress_to_zip": True,
            "telemetry_folder_history": [],
            "photos_folder_history": [],
            "relief_folder_history": [],
            "output_name_history": ["tele_photo.tlm"],
            "archive_template_history": ["Дата_{номер маршрута}"],
            "route_number": "M2.1",
            "selected_camera": "Ручная настройка",
            "com_port": "",
            "create_kml_files": True,
            "create_tab_files": True,
            "kml_opacity": "d6",
            "cameras": {
                "Ручная настройка": {
                    "focal_length": 50,
                    "sensor_width": 36,
                    "sensor_height": 24,
                    "rotate_raster": False,
                    "camera_rotation": 0
                },
                "Canon_500D_F55mm": {
                    "focal_length": 55,
                    "sensor_width": 22.3,
                    "sensor_height": 14.9,
                    "rotate_raster": False,
                    "camera_rotation": 0
                },
                "Samsung_PL10_F6.3mm": {
                    "focal_length": 6.3,
                    "sensor_width": 6.17,
                    "sensor_height": 4.55,
                    "rotate_raster": True,
                    "camera_rotation": 0
                },
                "Panasonic_DMC-GF1_F20mm": {
                    "focal_length": 20,
                    "sensor_width": 17.3,
                    "sensor_height": 13,
                    "rotate_raster": False,
                    "camera_rotation": 0
                },
                "FUJIFILM FinePIX XP30": {
                    "focal_length": 5,
                    "sensor_width": 6.17,
                    "sensor_height": 4.55,
                    "rotate_raster": True,
                    "camera_rotation": 0
                },
                "NIKON COOLPIX AW130": {
                    "focal_length": 4.5,
                    "sensor_width": 7.44,
                    "sensor_height": 5.58,
                    "rotate_raster": True,
                    "camera_rotation": 0
                },
                "Samsung Techwin <Samsung i8, Samsung VLUU i8>": {
                    "focal_length": 6.3,
                    "sensor_width": 7.81,
                    "sensor_height": 5.86,
                    "rotate_raster": True,
                    "camera_rotation": 0
                },
                "SONY RX1": {
                    "focal_length": 35,
                    "sensor_width": 35.8,
                    "sensor_height": 23.9,
                    "rotate_raster": False,
                    "camera_rotation": 0
                },
                "SONY Alpha 5100 (фокус20)": {
                    "focal_length": 20,
                    "sensor_width": 23.5,
                    "sensor_height": 15.6,
                    "rotate_raster": False,
                    "camera_rotation": 0
                },
                "SONY Alpha 5100 (фокус16)": {
                    "focal_length": 16,
                    "sensor_width": 23.5,
                    "sensor_height": 15.6,
                    "rotate_raster": False,
                    "camera_rotation": 0
                }
            }
        }
        
        # Объединяем с существующими настройками
        self.plugin_settings = {**default_plugin_settings, **plugin_settings}
        
    def save_plugin_settings(self):
        """Сохранение настроек плагина"""
        self.settings.settings["telemetry_plugin"] = self.plugin_settings
        self.settings.save_settings()
    
    def get_tab_name(self):
        return "Телеметрия фото"
    
    def create_tab(self):
        """Создание вкладки плагина с нотебуком"""
        tab_frame = ttk.Frame(self.root)
        
        # Создаем нотебук для разделения функционала
        notebook = ttk.Notebook(tab_frame)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка COM порта
        com_tab = self.create_com_tab()
        notebook.add(com_tab, text="COM порт")
        
        # Вкладка обработки телеметрии
        processing_tab = self.create_processing_tab()
        notebook.add(processing_tab, text="Обработка телеметрии")
        
        # Вкладка параметров фотоаппарата
        camera_tab = self.create_camera_tab()
        notebook.add(camera_tab, text="Параметры фотоаппарата")
        
        # Вкладка KML/TAB файлов
        kml_tab = self.create_kml_tab_tab()
        notebook.add(kml_tab, text="KML/TAB файлы")
        
        return tab_frame
    
    def create_com_tab(self):
        """Создание вкладки для работы с COM портом"""
        com_tab = ttk.Frame(self.root)
        
        main_frame = ttk.Frame(com_tab)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_label = ttk.Label(main_frame, text="Управление COM портом", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Фрейм для управления портами
        port_frame = ttk.LabelFrame(main_frame, text="Выбор COM порта")
        port_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Обнаружение портов
        ttk.Label(port_frame, text="Доступные COM порты:").pack(anchor=tk.W, pady=(5, 0))
        
        port_controls_frame = ttk.Frame(port_frame)
        port_controls_frame.pack(fill=tk.X, pady=5, padx=5)
        
        self.port_var = tk.StringVar(value=self.plugin_settings["com_port"])
        
        self.port_combo = ttk.Combobox(port_controls_frame, textvariable=self.port_var, width=15)
        self.port_combo.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(port_controls_frame, text="Обновить список", 
                  command=self.refresh_com_ports).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(port_controls_frame, text="Автоопределение", 
                  command=self.auto_detect_port).pack(side=tk.LEFT, padx=5)
        
        # Управление подключением
        connection_frame = ttk.Frame(port_frame)
        connection_frame.pack(fill=tk.X, pady=5, padx=5)
        
        self.connect_button = ttk.Button(
            connection_frame, 
            text="Открыть порт", 
            command=self.toggle_serial_connection,
            style="Accent.TButton"
        )
        self.connect_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(connection_frame, text="Считать телеметрию", 
                  command=self.read_telemetry).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(connection_frame, text="Остановить чтение", 
                  command=self.stop_reading_telemetry).pack(side=tk.LEFT, padx=5)
        
        # Статус подключения
        self.status_label = ttk.Label(port_frame, text="Статус: Порт не открыт", foreground="red")
        self.status_label.pack(anchor=tk.W, padx=5, pady=5)
        
        # Логи COM порта
        log_frame = ttk.LabelFrame(main_frame, text="Логи COM порта")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.com_log_text = tk.Text(log_frame, wrap=tk.WORD, height=10, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.com_log_text.yview)
        self.com_log_text.configure(yscrollcommand=scrollbar.set)
        
        self.com_log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки для логов COM порта
        com_log_buttons = ttk.Frame(main_frame)
        com_log_buttons.pack(fill=tk.X, pady=5)
        
        ttk.Button(com_log_buttons, text="Очистить логи COM", 
                  command=self.clear_com_logs).pack(side=tk.RIGHT, padx=5)
        
        # Обновляем список портов при создании
        self.refresh_com_ports()
        
        return com_tab
    
    def create_processing_tab(self):
        """Создание вкладки обработки телеметрии"""
        processing_tab = ttk.Frame(self.root)
        
        main_frame = ttk.Frame(processing_tab)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_label = ttk.Label(main_frame, text="Обработка телеметрии фотографий", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Фрейм для настроек
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки обработки")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Папка с телеметрией
        telemetry_frame = ttk.Frame(settings_frame)
        telemetry_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(telemetry_frame, text="Файл телеметрии (.tlm):").pack(anchor=tk.W)
        
        self.telemetry_var = tk.StringVar(value=self.plugin_settings["telemetry_folder"])
        
        telemetry_combo = ttk.Combobox(
            telemetry_frame, 
            textvariable=self.telemetry_var,
            values=self.plugin_settings["telemetry_folder_history"],
            width=50
        )
        telemetry_combo.pack(fill=tk.X, pady=2)
        
        telemetry_buttons_frame = ttk.Frame(telemetry_frame)
        telemetry_buttons_frame.pack(fill=tk.X)
        
        ttk.Button(telemetry_buttons_frame, text="Обзор", 
                  command=self.browse_telemetry_file).pack(side=tk.LEFT, pady=2)
        ttk.Button(telemetry_buttons_frame, text="Сканировать папку", 
                  command=self.scan_for_telemetry).pack(side=tk.LEFT, padx=5)
        
        # Папка с фотографиями
        photos_frame = ttk.Frame(settings_frame)
        photos_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(photos_frame, text="Папка с фотографиями:").pack(anchor=tk.W)
        
        self.photos_var = tk.StringVar(value=self.plugin_settings["photos_folder"])
        
        photos_combo = ttk.Combobox(
            photos_frame, 
            textvariable=self.photos_var,
            values=self.plugin_settings["photos_folder_history"],
            width=50
        )
        photos_combo.pack(fill=tk.X, pady=2)
        ttk.Button(photos_frame, text="Обзор", 
                  command=self.browse_photos_folder).pack(anchor=tk.W, pady=2)
        
        # Папка с рельефом
        relief_frame = ttk.Frame(settings_frame)
        relief_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(relief_frame, text="Папка с рельефом (HGT):").pack(anchor=tk.W)
        
        self.relief_var = tk.StringVar(value=self.plugin_settings["relief_folder"])
        
        relief_combo = ttk.Combobox(
            relief_frame, 
            textvariable=self.relief_var,
            values=self.plugin_settings["relief_folder_history"],
            width=50
        )
        relief_combo.pack(fill=tk.X, pady=2)
        ttk.Button(relief_frame, text="Обзор", 
                  command=self.browse_relief_folder).pack(anchor=tk.W, pady=2)
        
        # Настройки выходных файлов
        output_frame = ttk.Frame(settings_frame)
        output_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Имя выходного файла телеметрии
        ttk.Label(output_frame, text="Имя выходного файла телеметрии:").pack(anchor=tk.W)
        
        self.output_name_var = tk.StringVar(value=self.plugin_settings["output_telemetry_name"])
        
        output_name_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.output_name_var,
            values=self.plugin_settings["output_name_history"],
            width=30
        )
        output_name_combo.pack(fill=tk.X, pady=2)
        
        # Шаблон архива
        ttk.Label(output_frame, text="Шаблон имени архива:").pack(anchor=tk.W, pady=(10, 0))
        
        self.archive_var = tk.StringVar(value=self.plugin_settings["archive_template"])
        
        archive_combo = ttk.Combobox(
            output_frame, 
            textvariable=self.archive_var,
            values=self.plugin_settings["archive_template_history"],
            width=30
        )
        archive_combo.pack(fill=tk.X, pady=2)
        
        # Номер маршрута
        route_frame = ttk.Frame(output_frame)
        route_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(route_frame, text="Номер маршрута:").pack(side=tk.LEFT)
        
        self.route_var = tk.StringVar(value=self.plugin_settings["route_number"])
        
        route_entry = ttk.Entry(route_frame, textvariable=self.route_var, width=15)
        route_entry.pack(side=tk.LEFT, padx=5)
        
        # Опции
        options_frame = ttk.Frame(settings_frame)
        options_frame.pack(fill=tk.X, pady=5, padx=5)
        
        self.compress_var = tk.BooleanVar(value=self.plugin_settings["compress_to_zip"])
        ttk.Checkbutton(options_frame, text="Создать ZIP архив после обработки", 
                       variable=self.compress_var).pack(anchor=tk.W)
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="Сохранить настройки", 
                  command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Обработать телеметрию", 
                  command=self.process_telemetry, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Создать конфиг для Программа2", 
                  command=self.create_program2_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Очистить логи", 
                  command=self.clear_logs).pack(side=tk.RIGHT, padx=5)
        
        # Логи
        log_frame = ttk.LabelFrame(main_frame, text="Логи обработки")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=15, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        return processing_tab
    
    def create_camera_tab(self):
        """Создание вкладки параметров фотоаппарата"""
        camera_tab = ttk.Frame(self.root)
        
        main_frame = ttk.Frame(camera_tab)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_label = ttk.Label(main_frame, text="Параметры фотоаппарата", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Выбор фотоаппарата
        selection_frame = ttk.LabelFrame(main_frame, text="Выбор фотоаппарата")
        selection_frame.pack(fill=tk.X, pady=(0, 10))
        
        selection_controls_frame = ttk.Frame(selection_frame)
        selection_controls_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(selection_controls_frame, text="Модель:").pack(side=tk.LEFT)
        
        self.camera_var = tk.StringVar(value=self.plugin_settings["selected_camera"])
        self.camera_var.trace('w', self.on_camera_selected)
        
        camera_combo = ttk.Combobox(
            selection_controls_frame, 
            textvariable=self.camera_var,
            values=list(self.plugin_settings["cameras"].keys()),
            width=30,
            state="readonly"
        )
        camera_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(selection_controls_frame, text="Добавить модель", 
                  command=self.add_camera_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(selection_controls_frame, text="Удалить модель", 
                  command=self.delete_camera).pack(side=tk.LEFT, padx=5)
        
        # Параметры фотоаппарата
        params_frame = ttk.LabelFrame(main_frame, text="Параметры выбранного фотоаппарата")
        params_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Фокусное расстояние
        focal_frame = ttk.Frame(params_frame)
        focal_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(focal_frame, text="Фокусное расстояние (мм):", width=25).pack(side=tk.LEFT)
        self.focal_var = tk.DoubleVar()
        focal_entry = ttk.Entry(focal_frame, textvariable=self.focal_var, width=15)
        focal_entry.pack(side=tk.LEFT)
        
        # Ширина матрицы
        width_frame = ttk.Frame(params_frame)
        width_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(width_frame, text="Ширина матрицы (мм):", width=25).pack(side=tk.LEFT)
        self.sensor_width_var = tk.DoubleVar()
        width_entry = ttk.Entry(width_frame, textvariable=self.sensor_width_var, width=15)
        width_entry.pack(side=tk.LEFT)
        
        # Высота матрицы
        height_frame = ttk.Frame(params_frame)
        height_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(height_frame, text="Высота матрицы (мм):", width=25).pack(side=tk.LEFT)
        self.sensor_height_var = tk.DoubleVar()
        height_entry = ttk.Entry(height_frame, textvariable=self.sensor_height_var, width=15)
        height_entry.pack(side=tk.LEFT)
        
        # Поворот растра
        rotate_frame = ttk.Frame(params_frame)
        rotate_frame.pack(fill=tk.X, pady=5, padx=5)
        
        self.rotate_var = tk.BooleanVar()
        ttk.Checkbutton(rotate_frame, text="Поворот растра", 
                       variable=self.rotate_var).pack(side=tk.LEFT)
        
        # Поворот камеры
        rotation_frame = ttk.Frame(params_frame)
        rotation_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(rotation_frame, text="Поворот камеры (градусы):", width=25).pack(side=tk.LEFT)
        self.camera_rotation_var = tk.DoubleVar()
        rotation_entry = ttk.Entry(rotation_frame, textvariable=self.camera_rotation_var, width=15)
        rotation_entry.pack(side=tk.LEFT)
        
        # Кнопки управления
        camera_buttons_frame = ttk.Frame(main_frame)
        camera_buttons_frame.pack(fill=tk.X)
        
        ttk.Button(camera_buttons_frame, text="Сохранить параметры", 
                  command=self.save_camera_params).pack(side=tk.LEFT, padx=5)
        ttk.Button(camera_buttons_frame, text="Сбросить к значениям по умолчанию", 
                  command=self.reset_camera_params).pack(side=tk.LEFT, padx=5)
        
        # Загружаем параметры выбранного фотоаппарата
        self.load_camera_params()
        
        return camera_tab
    
    def create_kml_tab_tab(self):
        """Создание вкладки для настройки KML/TAB файлов"""
        kml_tab = ttk.Frame(self.root)
        
        main_frame = ttk.Frame(kml_tab)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        title_label = ttk.Label(main_frame, text="Настройка KML и TAB файлов", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Фрейм для настроек KML/TAB
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки создания файлов")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Опции создания файлов
        options_frame = ttk.Frame(settings_frame)
        options_frame.pack(fill=tk.X, pady=5, padx=5)
        
        self.create_kml_var = tk.BooleanVar(value=self.plugin_settings.get("create_kml_files", True))
        ttk.Checkbutton(options_frame, text="Создавать KML файлы", 
                       variable=self.create_kml_var).pack(anchor=tk.W)
        
        self.create_tab_var = tk.BooleanVar(value=self.plugin_settings.get("create_tab_files", True))
        ttk.Checkbutton(options_frame, text="Создавать TAB файлы", 
                       variable=self.create_tab_var).pack(anchor=tk.W)
        
        # Настройки KML
        kml_settings_frame = ttk.Frame(settings_frame)
        kml_settings_frame.pack(fill=tk.X, pady=5, padx=5)
        
        ttk.Label(kml_settings_frame, text="Прозрачность KML (hex):").pack(side=tk.LEFT)
        
        self.kml_opacity_var = tk.StringVar(value=self.plugin_settings.get("kml_opacity", "d6"))
        opacity_entry = ttk.Entry(kml_settings_frame, textvariable=self.kml_opacity_var, width=5)
        opacity_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(kml_settings_frame, text="от 00 (прозрачный) до ff (непрозрачный)").pack(side=tk.LEFT)
        
        # Информация о файлах
        info_frame = ttk.LabelFrame(main_frame, text="Информация о создаваемых файлах")
        info_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        info_text = tk.Text(info_frame, wrap=tk.WORD, height=8, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(info_frame, orient=tk.VERTICAL, command=info_text.yview)
        info_text.configure(yscrollcommand=scrollbar.set)
        
        info_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Заполняем информацию
        info_text.config(state=tk.NORMAL)
        info_text.insert(tk.END, "KML файлы:\n")
        info_text.insert(tk.END, "• Создаются для каждой фотографии\n")
        info_text.insert(tk.END, "• Содержат геопривязку и метаданные\n")
        info_text.insert(tk.END, "• Могут быть открыты в Google Earth\n")
        info_text.insert(tk.END, "• Включают координаты углов, высоту, углы наклона\n\n")
        
        info_text.insert(tk.END, "TAB файлы:\n")
        info_text.insert(tk.END, "• Создаются для каждой фотографии\n")
        info_text.insert(tk.END, "• Используются в MapInfo\n")
        info_text.insert(tk.END, "• Содержат координаты углов растрового изображения\n")
        info_text.config(state=tk.DISABLED)
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Сохранить настройки", 
                  command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Просмотреть пример KML", 
                  command=self.show_kml_example).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Просмотреть пример TAB", 
                  command=self.show_tab_example).pack(side=tk.LEFT, padx=5)
        
        return kml_tab
    
    def refresh_com_ports(self):
        """Обновление списка COM портов"""
        ports = serial.tools.list_ports.comports()
        port_list = [port.device for port in ports]
        self.port_combo['values'] = port_list
        self.log_com_message(f"Найдено COM портов: {len(port_list)}")
    
    def auto_detect_port(self):
        """Автоматическое определение нового COM порта"""
        initial_ports = set([port.device for port in serial.tools.list_ports.comports()])
        
        self.log_com_message("Отключите и подключите фотоаппарат...")
        messagebox.showinfo("Автоопределение", 
                           "1. Отключите фотоаппарат от компьютера\n"
                           "2. Нажмите OK\n"
                           "3. Подключите фотоаппарат\n"
                           "4. Нажмите OK еще раз")
        
        final_ports = set([port.device for port in serial.tools.list_ports.comports()])
        new_ports = final_ports - initial_ports
        
        if new_ports:
            new_port = list(new_ports)[0]
            self.port_var.set(new_port)
            self.log_com_message(f"Обнаружен новый порт: {new_port}")
        else:
            self.log_com_message("Новые порты не обнаружены", "warning")
    
    def toggle_serial_connection(self):
        """Открытие/закрытие COM порта"""
        if self.serial_connection and self.serial_connection.is_open:
            self.close_serial_connection()
        else:
            self.open_serial_connection()
    
    def open_serial_connection(self):
        """Открытие COM порта"""
        port = self.port_var.get()
        if not port:
            messagebox.showerror("Ошибка", "Выберите COM порт")
            return
        
        try:
            self.serial_connection = serial.Serial(
                port=port,
                baudrate=9600,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1
            )
            
            self.connect_button.config(text="Закрыть порт")
            self.status_label.config(text=f"Статус: Порт {port} открыт", foreground="green")
            self.log_com_message(f"Порт {port} успешно открыт")
            
            # Сохраняем настройки
            self.plugin_settings["com_port"] = port
            self.save_plugin_settings()
            
        except Exception as e:
            self.log_com_message(f"Ошибка открытия порта {port}: {e}", "error")
            messagebox.showerror("Ошибка", f"Не удалось открыть порт {port}: {e}")
    
    def close_serial_connection(self):
        """Закрытие COM порта"""
        if self.serial_connection:
            try:
                self.serial_connection.close()
                self.connect_button.config(text="Открыть порт")
                self.status_label.config(text="Статус: Порт не открыт", foreground="red")
                self.log_com_message("COM порт закрыт")
            except Exception as e:
                self.log_com_message(f"Ошибка закрытия порта: {e}", "error")
    
    def read_telemetry(self):
        """Чтение телеметрии с COM порта"""
        if not self.serial_connection or not self.serial_connection.is_open:
            messagebox.showerror("Ошибка", "COM порт не открыт")
            return
        
        # Запрашиваем путь для сохранения файла телеметрии
        file_path = filedialog.asksaveasfilename(
            title="Сохранить телеметрию как",
            defaultextension=".tlm",
            filetypes=[("TLM files", "*.tlm"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        # Запускаем чтение в отдельном потоке
        self.is_reading_telemetry = True
        thread = threading.Thread(target=self._read_telemetry_thread, args=(file_path,))
        thread.daemon = True
        thread.start()
    
    def stop_reading_telemetry(self):
        """Остановка чтения телеметрии"""
        self.is_reading_telemetry = False
        self.log_com_message("Чтение телеметрии остановлено")
    
    def _read_telemetry_thread(self, file_path):
        """Поток чтения телеметрии"""
        try:
            self.log_com_message("Начало чтения телеметрии...")
            
            with open(file_path, 'w', encoding='utf-8') as f:
                while self.is_reading_telemetry and self.serial_connection and self.serial_connection.is_open:
                    if self.serial_connection.in_waiting > 0:
                        line = self.serial_connection.readline().decode('utf-8', errors='ignore').strip()
                        if line:
                            f.write(line + '\n')
                            f.flush()
                            self.log_com_message(f"Получено: {line}")
                    
                    # Небольшая задержка для снижения нагрузки на CPU
                    threading.Event().wait(0.1)
            
            self.log_com_message("Чтение телеметрии завершено")
            
        except Exception as e:
            self.log_com_message(f"Ошибка чтения телеметрии: {e}", "error")
    
    def browse_telemetry_file(self):
        """Выбор файла телеметрии"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл телеметрии",
            filetypes=[("TLM files", "*.tlm"), ("All files", "*.*")]
        )
        if file_path:
            self.telemetry_var.set(file_path)
            self.add_to_history("telemetry_folder_history", file_path)
    
    def browse_photos_folder(self):
        """Выбор папки с фотографиями"""
        folder = filedialog.askdirectory(title="Выберите папку с фотографиями")
        if folder:
            self.photos_var.set(folder)
            self.add_to_history("photos_folder_history", folder)
    
    def browse_relief_folder(self):
        """Выбор папки с рельефом"""
        folder = filedialog.askdirectory(title="Выберите папку с рельефом (HGT)")
        if folder:
            self.relief_var.set(folder)
            self.add_to_history("relief_folder_history", folder)
    
    def scan_for_telemetry(self):
        """Сканирование папки на наличие файлов телеметрии"""
        folder = filedialog.askdirectory(title="Выберите папку для сканирования")
        if folder:
            tlm_files = list(Path(folder).glob("*.tlm"))
            if tlm_files:
                # Обновляем историю
                for tlm_file in tlm_files:
                    self.add_to_history("telemetry_folder_history", str(tlm_file))
                
                # Устанавливаем первый найденный файл
                self.telemetry_var.set(str(tlm_files[0]))
                self.log_message(f"Найдено файлов .tlm: {len(tlm_files)}")
            else:
                messagebox.showwarning("Внимание", "В выбранной папке не найдено файлов .tlm")
    
    def add_to_history(self, history_key, value):
        """Добавление значения в историю"""
        if value and value not in self.plugin_settings[history_key]:
            self.plugin_settings[history_key].insert(0, value)
            # Ограничиваем историю 10 элементами
            self.plugin_settings[history_key] = self.plugin_settings[history_key][:10]
            self.save_plugin_settings()
    
    def on_camera_selected(self, *args):
        """Обработка выбора фотоаппарата"""
        self.load_camera_params()
    
    def load_camera_params(self):
        """Загрузка параметров выбранного фотоаппарата"""
        camera_name = self.camera_var.get()
        if camera_name in self.plugin_settings["cameras"]:
            camera_params = self.plugin_settings["cameras"][camera_name]
            
            self.focal_var.set(camera_params["focal_length"])
            self.sensor_width_var.set(camera_params["sensor_width"])
            self.sensor_height_var.set(camera_params["sensor_height"])
            self.rotate_var.set(camera_params["rotate_raster"])
            self.camera_rotation_var.set(camera_params["camera_rotation"])
    
    def save_camera_params(self):
        """Сохранение параметров фотоаппарата"""
        camera_name = self.camera_var.get()
        
        if camera_name in self.plugin_settings["cameras"]:
            self.plugin_settings["cameras"][camera_name] = {
                "focal_length": self.focal_var.get(),
                "sensor_width": self.sensor_width_var.get(),
                "sensor_height": self.sensor_height_var.get(),
                "rotate_raster": self.rotate_var.get(),
                "camera_rotation": self.camera_rotation_var.get()
            }
            
            self.plugin_settings["selected_camera"] = camera_name
            self.save_plugin_settings()
            self.log_message(f"Параметры фотоаппарата '{camera_name}' сохранены")
        else:
            self.log_message("Ошибка: фотоаппарат не найден", "error")
    
    def add_camera_dialog(self):
        """Диалог добавления нового фотоаппарата"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Добавить новый фотоаппарат")
        dialog.geometry("300x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Название фотоаппарата:").pack(anchor=tk.W, pady=(0, 5))
        
        name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=name_var, width=30)
        name_entry.pack(fill=tk.X, pady=(0, 15))
        
        def add_camera():
            name = name_var.get().strip()
            if not name:
                messagebox.showerror("Ошибка", "Введите название фотоаппарата")
                return
            
            if name in self.plugin_settings["cameras"]:
                messagebox.showerror("Ошибка", "Фотоаппарат с таким названием уже существует")
                return
            
            # Добавляем новый фотоаппарат с параметрами по умолчанию
            self.plugin_settings["cameras"][name] = {
                "focal_length": 50,
                "sensor_width": 36,
                "sensor_height": 24,
                "rotate_raster": False,
                "camera_rotation": 0
            }
            
            self.save_plugin_settings()
            
            # Обновляем комбобокс
            self.camera_var.set(name)
            self.load_camera_params()
            
            dialog.destroy()
            self.log_message(f"Добавлен новый фотоаппарат: {name}")
        
        ttk.Button(main_frame, text="Добавить", command=add_camera).pack(pady=5)
    
    def delete_camera(self):
        """Удаление выбранного фотоаппарата"""
        camera_name = self.camera_var.get()
        
        if camera_name == "Ручная настройка":
            messagebox.showerror("Ошибка", "Нельзя удалить базовый профиль 'Ручная настройка'")
            return
        
        if messagebox.askyesno("Подтверждение", f"Удалить фотоаппарат '{camera_name}'?"):
            del self.plugin_settings["cameras"][camera_name]
            self.plugin_settings["selected_camera"] = "Ручная настройка"
            self.save_plugin_settings()
            
            # Обновляем комбобокс
            self.camera_var.set("Ручная настройка")
            self.load_camera_params()
            
            self.log_message(f"Фотоаппарат '{camera_name}' удален")
    
    def reset_camera_params(self):
        """Сброс параметров к значениям по умолчанию"""
        camera_name = self.camera_var.get()
        if camera_name in self.plugin_settings["cameras"]:
            # Для ручной настройки сбрасываем к базовым значениям
            if camera_name == "Ручная настройка":
                default_params = {
                    "focal_length": 50,
                    "sensor_width": 36,
                    "sensor_height": 24,
                    "rotate_raster": False,
                    "camera_rotation": 0
                }
            else:
                # Для других моделей загружаем сохраненные значения
                default_params = self.plugin_settings["cameras"][camera_name]
            
            self.focal_var.set(default_params["focal_length"])
            self.sensor_width_var.set(default_params["sensor_width"])
            self.sensor_height_var.set(default_params["sensor_height"])
            self.rotate_var.set(default_params["rotate_raster"])
            self.camera_rotation_var.set(default_params["camera_rotation"])
            
            self.log_message(f"Параметры фотоаппарата '{camera_name}' сброшены")
    
    def show_kml_example(self):
        """Показать пример KML файла"""
        example_kml = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2">
<GroundOverlay>
    <name>DSC02537.JPG</name>
    <color>d6ffffff</color>
    <Icon>
        <href>DSC02537.JPG</href>
        <viewBoundScale>0.2</viewBoundScale>
    </Icon>
    <gx:LatLonQuad>
        <coordinates>
            51.087966664826304,53.26976848712158 
            51.082268173112475,53.27182116303870 
            51.079995355178724,53.26955337583642 
            51.085693654229381,53.26750080774580
        </coordinates>
    </gx:LatLonQuad>
    <ExtendedData>
        <TimeStamp>
            <when>2025/10/01T12:08:07+00:00</when>
        </TimeStamp>
        <Data name="FlightTimeOffset">
            <value>5458</value>
        </Data>
        <Data name="Speed">
            <value>95.00000</value>
        </Data>
        <Data name="RelativeAltitude">
            <value>598.40000</value>
        </Data>
        <Data name="Elevation">
            <value>660.60000</value>
        </Data>
        <Data name="Pitch">
            <value>-22.80000</value>
        </Data>
        <Data name="Roll">
            <value>-0.50000</value>
        </Data>
        <Data name="Yaw">
            <value>211.00000</value>
        </Data>
    </ExtendedData>
</GroundOverlay>
</kml>"""
        
        self.show_text_dialog("Пример KML файла", example_kml)
    
    def show_tab_example(self):
        """Показать пример TAB файла"""
        example_tab = """!table
!version 300
!charset WindowsLatin1

Definition Table
File "DSC02537.JPG"
Type "RASTER"
(51.08569365,53.26750081) (0,0) Label "TL",
(51.07999536,53.26955338) (6000,0) Label "TR",
(51.08226817,53.27182116) (6000,4000) Label "BR",
(51.08796666,53.26976849) (0,4000) Label "BL"
CoordSys Earth Projection 1, 104
Units "degree" """
        
        self.show_text_dialog("Пример TAB файла", example_tab)
    
    def show_text_dialog(self, title, content):
        """Показать диалог с текстом"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        text_widget = tk.Text(main_frame, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        text_widget.insert(tk.END, content)
        text_widget.config(state=tk.DISABLED)
        
        ttk.Button(main_frame, text="Закрыть", command=dialog.destroy).pack(pady=10)
    
    def calculate_image_corners(self, center_lat, center_lon, altitude, yaw, pitch, roll, focal_length, sensor_width, sensor_height):
        """
        Расчет координат углов фотографии на основе параметров камеры и положения
        """
        try:
            # Преобразуем углы в радианы
            yaw_rad = math.radians(yaw)
            pitch_rad = math.radians(pitch)
            roll_rad = math.radians(roll)
            
            # Расчет углов обзора
            fov_x = 2 * math.atan(sensor_width / (2 * focal_length))
            fov_y = 2 * math.atan(sensor_height / (2 * focal_length))
            
            # Расчет размеров покрытия на земле
            ground_width = 2 * altitude * math.tan(fov_x / 2)
            ground_height = 2 * altitude * math.tan(fov_y / 2)
            
            # Углы поворота для каждого угла изображения
            corners = [
                (-ground_width/2, -ground_height/2),  # Нижний левый
                (ground_width/2, -ground_height/2),   # Нижний правый
                (ground_width/2, ground_height/2),    # Верхний правый
                (-ground_width/2, ground_height/2)    # Верхний левый
            ]
            
            # Применяем поворот
            rotated_corners = []
            for dx, dy in corners:
                # Поворот по рысканью (yaw)
                x_rot = dx * math.cos(yaw_rad) - dy * math.sin(yaw_rad)
                y_rot = dx * math.sin(yaw_rad) + dy * math.cos(yaw_rad)
                
                # Поворот по тангажу (pitch) и крену (roll) упрощенно
                x_final = x_rot
                y_final = y_rot
                
                rotated_corners.append((x_final, y_final))
            
            # Конвертируем смещения в координаты
            earth_radius = 6371000  # Радиус Земли в метрах
            
            corner_coords = []
            for dx, dy in rotated_corners:
                # Смещение в градусах широты
                dlat = (dy / earth_radius) * (180 / math.pi)
                # Смещение в градусах долготы (учитываем широту)
                dlon = (dx / (earth_radius * math.cos(math.radians(center_lat)))) * (180 / math.pi)
                
                corner_lat = center_lat + dlat
                corner_lon = center_lon + dlon
                
                corner_coords.append((corner_lon, corner_lat))
            
            return corner_coords
            
        except Exception as e:
            self.log_message(f"Ошибка расчета углов изображения: {e}", "error")
            # Возвращаем углы по умолчанию (квадрат вокруг центра)
            delta = 0.001
            return [
                (center_lon - delta, center_lat - delta),  # BL
                (center_lon + delta, center_lat - delta),  # BR  
                (center_lon + delta, center_lat + delta),  # TR
                (center_lon - delta, center_lat + delta)   # TL
            ]
    
    def create_kml_file(self, photo_file, telemetry_data, output_folder):
        """Создание KML файла для фотографии"""
        try:
            # Получаем параметры камеры
            camera_name = self.plugin_settings["selected_camera"]
            camera_params = self.plugin_settings["cameras"].get(camera_name, {})
            
            # Расчет координат углов
            center_lat = telemetry_data.get('latitude', 53.26966)
            center_lon = telemetry_data.get('longitude', 51.08398)
            altitude = telemetry_data.get('altitude', 598.4)
            yaw = telemetry_data.get('yaw', 211.0)
            pitch = telemetry_data.get('pitch', -22.8)
            roll = telemetry_data.get('roll', -0.5)
            focal_length = camera_params.get('focal_length', 50)
            sensor_width = camera_params.get('sensor_width', 36)
            sensor_height = camera_params.get('sensor_height', 24)
            
            corners = self.calculate_image_corners(
                center_lat, center_lon, altitude, yaw, pitch, roll,
                focal_length, sensor_width, sensor_height
            )
            
            # Форматируем координаты для KML
            coordinates_str = ""
            for lon, lat in corners:
                coordinates_str += f"            {lon:.15f},{lat:.15f} \n"
            
            # Создаем KML содержимое
            kml_content = f"""<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<GroundOverlay>
    <name>{photo_file.name}</name>
    <color>{self.kml_opacity_var.get()}ffffff</color>
    <Icon>
        <href>{photo_file.name}</href>
        <viewBoundScale>0.2</viewBoundScale>
    </Icon>
    <gx:LatLonQuad>
        <coordinates>
{coordinates_str}        </coordinates>
    </gx:LatLonQuad>
    <ExtendedData xmlns:geogr="http://geogr.stc-orion.ru">
        <TimeStamp>
            <when>{telemetry_data.get('timestamp', '2025/10/01T12:08:07+00:00')}</when>
        </TimeStamp>
        <geogr:FlightTimeOffset>{telemetry_data.get('flight_time_offset', 5458)}</geogr:FlightTimeOffset>
        <geogr:Speed>{telemetry_data.get('speed', 95.0):.5f}</geogr:Speed>
        <geogr:RelativeAltitude>{telemetry_data.get('relative_altitude', 598.4):.5f}</geogr:RelativeAltitude>
        <geogr:Elevation>{telemetry_data.get('elevation', 660.6):.5f}</geogr:Elevation>
        <geogr:EulerAngles>
            <pitch>{pitch:.5f}</pitch>
            <roll>{roll:.5f}</roll>
            <yaw>{yaw:.5f}</yaw>
        </geogr:EulerAngles>
        <geogr:Location>
            <latitude>{center_lat:.5f}</latitude>
            <longitude>{center_lon:.5f}</longitude>
        </geogr:Location>
    </ExtendedData>
</GroundOverlay>
</kml>"""
            
            # Сохраняем файл
            kml_file = Path(output_folder) / f"{photo_file.stem}.kml"
            with open(kml_file, 'w', encoding='utf-8') as f:
                f.write(kml_content)
            
            self.log_message(f"Создан KML файл: {kml_file.name}")
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка создания KML файла для {photo_file.name}: {e}", "error")
            return False
    
    def create_tab_file(self, photo_file, telemetry_data, output_folder):
        """Создание TAB файла для фотографии"""
        try:
            # Расчет координат углов (используем ту же логику, что и для KML)
            camera_name = self.plugin_settings["selected_camera"]
            camera_params = self.plugin_settings["cameras"].get(camera_name, {})
            
            center_lat = telemetry_data.get('latitude', 53.26966)
            center_lon = telemetry_data.get('longitude', 51.08398)
            altitude = telemetry_data.get('altitude', 598.4)
            yaw = telemetry_data.get('yaw', 211.0)
            pitch = telemetry_data.get('pitch', -22.8)
            roll = telemetry_data.get('roll', -0.5)
            focal_length = camera_params.get('focal_length', 50)
            sensor_width = camera_params.get('sensor_width', 36)
            sensor_height = camera_params.get('sensor_height', 24)
            
            corners = self.calculate_image_corners(
                center_lat, center_lon, altitude, yaw, pitch, roll,
                focal_length, sensor_width, sensor_height
            )
            
            # Для TAB файла порядок углов: BL, BR, TR, TL
            bl_lon, bl_lat = corners[0]  # Нижний левый
            br_lon, br_lat = corners[1]  # Нижний правый
            tr_lon, tr_lat = corners[2]  # Верхний правый
            tl_lon, tl_lat = corners[3]  # Верхний левый
            
            # Создаем TAB содержимое
            tab_content = f"""!table
!version 300
!charset WindowsLatin1

Definition Table
File "{photo_file.name}"
Type "RASTER"
({bl_lon:.8f},{bl_lat:.8f}) (0,0) Label "BL",
({br_lon:.8f},{br_lat:.8f}) (6000,0) Label "BR",
({tr_lon:.8f},{tr_lat:.8f}) (6000,4000) Label "TR",
({tl_lon:.8f},{tl_lat:.8f}) (0,4000) Label "TL"
CoordSys Earth Projection 1, 104
Units "degree"
"""
            
            # Сохраняем файл
            tab_file = Path(output_folder) / f"{photo_file.stem}.tab"
            with open(tab_file, 'w', encoding='utf-8') as f:
                f.write(tab_content)
            
            self.log_message(f"Создан TAB файл: {tab_file.name}")
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка создания TAB файла для {photo_file.name}: {e}", "error")
            return False
    
    def get_exif_datetime(self, image_path):
        """Получение даты и времени из EXIF данных фотографии"""
        try:
            with Image.open(image_path) as img:
                exif_data = img._getexif()
                if exif_data:
                    for tag_id, value in exif_data.items():
                        tag = TAGS.get(tag_id, tag_id)
                        if tag == 'DateTime':
                            # Формат: "2023:10:09 09:46:16"
                            dt_str = value
                            dt = datetime.strptime(dt_str, '%Y:%m:%d %H:%M:%S')
                            return dt.strftime('%Y/%m/%d'), dt.strftime('%H:%M:%S'), dt
        except Exception as e:
            self.log_message(f"Ошибка чтения EXIF {image_path}: {e}", "warning")
        
        # Если EXIF нет, используем время создания файла
        try:
            file_time = datetime.fromtimestamp(os.path.getctime(image_path))
            return file_time.strftime('%Y/%m/%d'), file_time.strftime('%H:%M:%S'), file_time
        except:
            dt = datetime.now()
            return "0000/00/00", "00:00:00", dt
    
    def parse_telemetry_file(self, file_path):
        """Парсинг файла телеметрии с извлечением дополнительных данных"""
        telemetry_data = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    if line.startswith('L '):
                        parts = line.split()
                        if len(parts) >= 3:
                            # Парсим дату и время из телеметрии
                            date_str = parts[1]  # YYMMDD
                            time_str = parts[2]  # HHMMSS
                            
                            # Преобразуем в datetime
                            try:
                                dt = datetime.strptime(f"20{date_str} {time_str}", '%Y%m%d %H%M%S')
                                
                                # Создаем расширенную запись телеметрии
                                record = {
                                    'datetime': dt,
                                    'line': line,
                                    'line_num': line_num,
                                    'timestamp': dt.strftime('%Y/%m/%dT%H:%M:%S+00:00'),
                                    'flight_time_offset': line_num * 100,
                                    'speed': 95.0,
                                    'relative_altitude': 598.4,
                                    'elevation': 660.6,
                                    'pitch': -22.8,
                                    'roll': -0.5,
                                    'yaw': 211.0,
                                    'latitude': 53.26966,
                                    'longitude': 51.08398,
                                    'altitude': 598.4
                                }
                                
                                # Пытаемся извлечь реальные данные из строки телеметрии
                                if len(parts) > 10:
                                    try:
                                        # Пример парсинга реальных данных
                                        record['latitude'] = float(parts[3]) / 1000000 if len(parts) > 3 else 53.26966
                                        record['longitude'] = float(parts[4]) / 1000000 if len(parts) > 4 else 51.08398
                                        record['altitude'] = float(parts[5]) if len(parts) > 5 else 598.4
                                    except:
                                        pass
                                
                                telemetry_data.append(record)
                                
                            except ValueError as e:
                                self.log_message(f"Ошибка парсинга даты в строке {line_num}: {e}", "warning")
                        else:
                            self.log_message(f"Неверный формат строки {line_num}", "warning")
            
            self.log_message(f"Загружено записей телеметрии: {len(telemetry_data)}")
            return telemetry_data
            
        except Exception as e:
            self.log_message(f"Ошибка чтения файла телеметрии: {e}", "error")
            return []
    
    def find_closest_telemetry(self, photo_datetime, telemetry_data):
        """Поиск ближайшей записи телеметрии к времени фотографии"""
        if not telemetry_data:
            return None
        
        min_diff = None
        closest_record = None
        
        for record in telemetry_data:
            diff = abs((record['datetime'] - photo_datetime).total_seconds())
            if min_diff is None or diff < min_diff:
                min_diff = diff
                closest_record = record
        
        return closest_record
    
    def process_telemetry(self):
        """Основной процесс обработки телеметрии"""
        # Проверяем настройки
        telemetry_file = self.telemetry_var.get()
        photos_folder = self.photos_var.get()
        output_name = self.output_name_var.get()
        
        if not telemetry_file or not os.path.exists(telemetry_file):
            messagebox.showerror("Ошибка", "Укажите корректный файл телеметрии")
            return
        
        if not photos_folder or not os.path.exists(photos_folder):
            messagebox.showerror("Ошибка", "Укажите корректную папку с фотографиями")
            return
        
        # Запускаем в отдельном потоке
        thread = threading.Thread(target=self._process_telemetry_thread, 
                                 args=(telemetry_file, photos_folder, output_name))
        thread.daemon = True
        thread.start()
    
    def _process_telemetry_thread(self, telemetry_file, photos_folder, output_name):
        """Поток обработки телеметрии с созданием KML/TAB файлов"""
        try:
            self.log_message("Начало обработки телеметрии...")
            
            # Парсим телеметрию
            telemetry_data = self.parse_telemetry_file(telemetry_file)
            if not telemetry_data:
                self.log_message("Нет данных телеметрии для обработки", "error")
                return
            
            # Получаем список фотографий
            photo_extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG']
            photo_files = []
            for ext in photo_extensions:
                photo_files.extend(Path(photos_folder).glob(f"*{ext}"))
            
            self.log_message(f"Найдено фотографий: {len(photo_files)}")
            
            if not photo_files:
                self.log_message("В папке не найдено фотографий", "error")
                return
            
            # Создаем выходной файл
            output_path = Path(photos_folder) / output_name
            processed_count = 0
            
            with open(output_path, 'w', encoding='utf-8') as out_file:
                for photo_file in photo_files:
                    try:
                        # Получаем дату и время из EXIF
                        date_str, time_str, photo_datetime = self.get_exif_datetime(photo_file)
                        
                        # Ищем ближайшую запись телеметрии
                        closest_telemetry = self.find_closest_telemetry(photo_datetime, telemetry_data)
                        
                        if closest_telemetry:
                            # Формируем строку для выходного файла
                            telemetry_parts = closest_telemetry['line'].split()[3:]
                            telemetry_str = ' '.join(telemetry_parts[:10])
                            
                            output_line = f"{photo_file.name}\t{date_str}\t{time_str}\t{telemetry_str}\n"
                            out_file.write(output_line)
                            processed_count += 1
                            
                            # СОЗДАЕМ KML И TAB ФАЙЛЫ
                            if self.create_kml_var.get():
                                self.create_kml_file(photo_file, closest_telemetry, photos_folder)
                            
                            if self.create_tab_var.get():
                                self.create_tab_file(photo_file, closest_telemetry, photos_folder)
                            
                            self.log_message(f"Обработано: {photo_file.name} -> {closest_telemetry['line_num']}")
                        else:
                            self.log_message(f"Не найдена телеметрия для {photo_file.name}", "warning")
                            
                    except Exception as e:
                        self.log_message(f"Ошибка обработки {photo_file.name}: {e}", "error")
                        continue
            
            self.log_message(f"Обработка завершена. Обработано фотографий: {processed_count}/{len(photo_files)}")
            
            # Создаем архив если нужно
            if self.compress_var.get():
                self.create_archive(photos_folder, processed_count)
                
        except Exception as e:
            self.log_message(f"Критическая ошибка обработки: {e}", "error")
    
    def create_archive(self, photos_folder, file_count):
        """Создание ZIP архива"""
        try:
            # Формируем имя архива по шаблону
            archive_template = self.archive_var.get()
            route_number = self.route_var.get()
            current_date = datetime.now().strftime("%Y%m%d")
            
            archive_name = archive_template.replace("{номер маршрута}", route_number)
            archive_name = archive_name.replace("{дата}", current_date)
            archive_name = archive_name.replace("{date}", current_date)
            
            # Добавляем расширение если нужно
            if not archive_name.endswith('.zip'):
                archive_name += '.zip'
            
            archive_path = Path(photos_folder) / archive_name
            
            # Создаем архив
            with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Добавляем все фотографии и файлы
                for file_path in Path(photos_folder).iterdir():
                    if (file_path.suffix.lower() in ['.jpg', '.jpeg', '.png', '.tlm', '.kml', '.tab'] and 
                        file_path.name != archive_name):
                        zipf.write(file_path, file_path.name)
            
            self.log_message(f"Создан архив: {archive_name} ({file_count} файлов)")
            
        except Exception as e:
            self.log_message(f"Ошибка создания архива: {e}", "error")
    
    def create_program2_config(self):
        """Создание конфигурационного файла для Программа2"""
        try:
            # Проверяем необходимые настройки
            relief_folder = self.relief_var.get()
            photos_folder = self.photos_var.get()
            telemetry_file = self.telemetry_var.get()
            
            if not relief_folder or not os.path.exists(relief_folder):
                messagebox.showerror("Ошибка", "Укажите корректную папку с рельефом")
                return
            
            if not photos_folder or not os.path.exists(photos_folder):
                messagebox.showerror("Ошибка", "Укажите корректную папку с фотографиями")
                return
            
            if not telemetry_file or not os.path.exists(telemetry_file):
                messagebox.showerror("Ошибка", "Укажите корректный файл телеметрии")
                return
            
            # Получаем параметры выбранного фотоаппарата
            camera_name = self.camera_var.get()
            camera_params = self.plugin_settings["cameras"].get(camera_name, {})
            
            # Создаем конфигурацию
            config = {
                "relief_folder": relief_folder,
                "telemetry_file": telemetry_file,
                "photos_folder": photos_folder,
                "camera_parameters": camera_params,
                "camera_model": camera_name,
                "created_at": datetime.now().isoformat()
            }
            
            # Сохраняем конфигурационный файл
            config_path = Path(photos_folder) / "program2_config.json"
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self.log_message(f"Создан конфигурационный файл: {config_path}")
            messagebox.showinfo("Успех", f"Конфигурационный файл создан:\n{config_path}")
            
        except Exception as e:
            self.log_message(f"Ошибка создания конфигурационного файла: {e}", "error")
            messagebox.showerror("Ошибка", f"Ошибка создания конфигурационного файла: {e}")
    
    def save_settings(self):
        """Сохранение настроек плагина"""
        try:
            # Сохраняем текущие значения
            self.plugin_settings["telemetry_folder"] = self.telemetry_var.get()
            self.plugin_settings["photos_folder"] = self.photos_var.get()
            self.plugin_settings["relief_folder"] = self.relief_var.get()
            self.plugin_settings["output_telemetry_name"] = self.output_name_var.get()
            self.plugin_settings["archive_template"] = self.archive_var.get()
            self.plugin_settings["compress_to_zip"] = self.compress_var.get()
            self.plugin_settings["route_number"] = self.route_var.get()
            self.plugin_settings["selected_camera"] = self.camera_var.get()
            self.plugin_settings["com_port"] = self.port_var.get()
            self.plugin_settings["create_kml_files"] = self.create_kml_var.get()
            self.plugin_settings["create_tab_files"] = self.create_tab_var.get()
            self.plugin_settings["kml_opacity"] = self.kml_opacity_var.get()
            
            # Добавляем в историю
            self.add_to_history("output_name_history", self.output_name_var.get())
            self.add_to_history("archive_template_history", self.archive_var.get())
            
            self.save_plugin_settings()
            self.log_message("Настройки сохранены успешно")
            messagebox.showinfo("Успех", "Настройки плагина сохранены!")
            
        except Exception as e:
            self.log_message(f"Ошибка сохранения настроек: {e}", "error")
            messagebox.showerror("Ошибка", f"Ошибка сохранения настроек: {e}")
    
    def log_message(self, message, level="info"):
        """Добавление сообщения в лог обработки"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        
        # Добавляем в очередь для безопасного обновления GUI
        self.log_queue.put((log_entry, level, "processing"))
    
    def log_com_message(self, message, level="info"):
        """Добавление сообщения в лог COM порта"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        
        # Добавляем в очередь для безопасного обновления GUI
        self.log_queue.put((log_entry, level, "com"))
    
    def process_log_queue(self):
        """Обработка очереди логов"""
        try:
            while True:
                log_entry, level, log_type = self.log_queue.get_nowait()
                
                if log_type == "com":
                    log_widget = self.com_log_text
                else:
                    log_widget = self.log_text
                
                log_widget.config(state=tk.NORMAL)
                
                # Определяем цвет в зависимости от уровня
                if level == "error":
                    log_widget.insert(tk.END, log_entry, "error")
                elif level == "warning":
                    log_widget.insert(tk.END, log_entry, "warning")
                else:
                    log_widget.insert(tk.END, log_entry, "info")
                
                # Автопрокрутка
                log_widget.see(tk.END)
                log_widget.config(state=tk.DISABLED)
                
        except queue.Empty:
            pass
        finally:
            # Планируем следующую проверку
            self.root.after(100, self.process_log_queue)
    
    def clear_logs(self):
        """Очистка логов обработки"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def clear_com_logs(self):
        """Очистка логов COM порта"""
        self.com_log_text.config(state=tk.NORMAL)
        self.com_log_text.delete(1.0, tk.END)
        self.com_log_text.config(state=tk.DISABLED)

def get_plugin_class():
    return TelemetryPlugin