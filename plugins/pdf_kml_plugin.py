import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import re
from pathlib import Path
import threading
from datetime import datetime
import logging
import math

# Проверяем наличие необходимых библиотек
try:
    import PyPDF2
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False
    logging.error("PyPDF2 не установлен. Установите: pip install PyPDF2")

try:
    import simplekml
    KML_SUPPORT = True
except ImportError:
    KML_SUPPORT = False
    logging.error("simplekml не установлен. Установите: pip install simplekml")

class PDFDecoderPlugin:
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.loaded_files = []
        self.coordinate_format = tk.StringVar(value="degrees")
        self.kml_data = None
        self.takeoff_landing_radius = tk.DoubleVar(value=0.05)  # Радиус по умолчанию 50 метров
        
    def get_tab_name(self):
        return "PDF → KML"
    
    def create_tab(self):
        tab_frame = ttk.Frame(self.root)
        self.create_interface(tab_frame)
        return tab_frame
    
    def create_interface(self, parent):
        # Основной контейнер
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Конвертер PDF представлений в KML", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Проверка зависимостей
        if not PDF_SUPPORT or not KML_SUPPORT:
            warning_frame = ttk.Frame(main_frame)
            warning_frame.pack(fill=tk.X, pady=(0, 10))
            
            warning_text = "ВНИМАНИЕ: Не все зависимости установлены!\n"
            if not PDF_SUPPORT:
                warning_text += "• PyPDF2 не установлен\n"
            if not KML_SUPPORT:
                warning_text += "• simplekml не установлен\n"
            warning_text += "Установите: pip install PyPDF2 simplekml"
            
            ttk.Label(warning_frame, text=warning_text, foreground="red", 
                     justify=tk.LEFT).pack(anchor=tk.W)
        
        # Фрейм загрузки файлов
        upload_frame = ttk.LabelFrame(main_frame, text="Загрузка PDF файлов")
        upload_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Кнопки загрузки
        button_frame = ttk.Frame(upload_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="Добавить PDF файлы", 
                  command=self.add_pdf_files, state="normal" if PDF_SUPPORT else "disabled").pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="Очистить список", 
                  command=self.clear_files).pack(side=tk.LEFT, padx=2)
        
        # Список загруженных файлов
        self.files_listbox = tk.Listbox(upload_frame, height=6)
        self.files_listbox.pack(fill=tk.X, padx=5, pady=5)
        
        # Фрейм настроек
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки формата")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Выбор формата координат
        format_frame = ttk.Frame(settings_frame)
        format_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(format_frame, text="Формат координат:").pack(side=tk.LEFT)
        
        formats = [
            ("Градусы (56.123456, 52.123456)", "degrees"),
            ("Градусы-минуты (56°07.408'N, 52°19.356'E)", "degrees_minutes"),
            ("Градусы-минуты-секунды (56°07'24.5\"N, 52°19'21.4\"E)", "degrees_minutes_seconds")
        ]
        
        format_subframe = ttk.Frame(format_frame)
        format_subframe.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))
        
        for text, value in formats:
            ttk.Radiobutton(format_subframe, text=text, value=value, 
                           variable=self.coordinate_format).pack(anchor=tk.W)
        
        # Настройка радиуса точек взлета/посадки
        radius_frame = ttk.Frame(settings_frame)
        radius_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(radius_frame, text="Радиус точек взлета/посадки (км):").pack(side=tk.LEFT)
        
        radius_entry = ttk.Entry(radius_frame, textvariable=self.takeoff_landing_radius, width=8)
        radius_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(radius_frame, text="(рекомендуется 0.05-0.1 км)").pack(side=tk.LEFT)
        
        # Фрейм обработки
        process_frame = ttk.LabelFrame(main_frame, text="Обработка и экспорт")
        process_frame.pack(fill=tk.X, pady=(0, 10))
        
        process_buttons = ttk.Frame(process_frame)
        process_buttons.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(process_buttons, text="Обработать файлы", 
                  command=self.process_files, state="normal" if PDF_SUPPORT and KML_SUPPORT else "disabled").pack(side=tk.LEFT, padx=2)
        ttk.Button(process_buttons, text="Экспорт KML", 
                  command=self.export_kml, state="normal" if KML_SUPPORT else "disabled").pack(side=tk.LEFT, padx=2)
        ttk.Button(process_buttons, text="Показать координаты", 
                  command=self.show_coordinates, state="normal" if PDF_SUPPORT else "disabled").pack(side=tk.LEFT, padx=2)
        
        # Область результатов
        result_frame = ttk.LabelFrame(main_frame, text="Результаты обработки")
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # Текстовое поле с прокруткой для результатов
        text_frame = ttk.Frame(result_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.result_text = tk.Text(text_frame, height=15, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, 
                                 command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Статус бар
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, style='TLabel')
        status_bar.pack(fill=tk.X, pady=(5, 0))
    
    def add_pdf_files(self):
        """Добавление PDF файлов"""
        if not PDF_SUPPORT:
            messagebox.showerror("Ошибка", "PyPDF2 не установлен. Установите: pip install PyPDF2")
            return
            
        files = filedialog.askopenfilenames(
            title="Выберите PDF файлы представлений",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        for file_path in files:
            if file_path not in self.loaded_files:
                self.loaded_files.append(file_path)
                filename = os.path.basename(file_path)
                self.files_listbox.insert(tk.END, filename)
        
        self.update_status(f"Загружено файлов: {len(self.loaded_files)}")
    
    def clear_files(self):
        """Очистка списка файлов"""
        self.loaded_files.clear()
        self.files_listbox.delete(0, tk.END)
        self.update_status("Список файлов очищен")
    
    def process_files(self):
        """Обработка PDF файлов в отдельном потоке"""
        if not self.loaded_files:
            messagebox.showwarning("Внимание", "Нет загруженных PDF файлов")
            return
            
        if not PDF_SUPPORT:
            messagebox.showerror("Ошибка", "PyPDF2 не установлен. Установите: pip install PyPDF2")
            return
        
        if not KML_SUPPORT:
            messagebox.showerror("Ошибка", "simplekml не установлен. Установите: pip install simplekml")
            return
        
        self.update_status("Обработка файлов...")
        self.result_text.delete(1.0, tk.END)
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self._process_files_thread)
        thread.daemon = True
        thread.start()
    
    def _process_files_thread(self):
        """Поток обработки файлов"""
        try:
            all_data = []
            
            for file_path in self.loaded_files:
                try:
                    file_data = self.parse_pdf_file(file_path)
                    if file_data:
                        all_data.append(file_data)
                        self.root.after(0, lambda f=os.path.basename(file_path): 
                                       self.result_text.insert(tk.END, f"✓ Обработан: {f}\n"))
                    else:
                        self.root.after(0, lambda f=os.path.basename(file_path): 
                                       self.result_text.insert(tk.END, f"✗ Ошибка: {f}\n"))
                except Exception as e:
                    error_msg = f"Ошибка обработки {os.path.basename(file_path)}: {str(e)}"
                    self.root.after(0, lambda msg=error_msg: self.result_text.insert(tk.END, msg + "\n"))
            
            # Создание KML
            if all_data:
                self.kml_data = self.create_kml_data(all_data)
                self.root.after(0, self._processing_complete)
            else:
                self.root.after(0, lambda: self.update_status("Нет данных для создания KML"))
                
        except Exception as e:
            self.root.after(0, lambda: self.update_status(f"Ошибка обработки: {str(e)}"))
    
    def parse_pdf_file(self, file_path):
        """Парсинг PDF файла"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                
                for page in pdf_reader.pages:
                    text += page.extract_text()
            
            return self.extract_data_from_text(text, os.path.basename(file_path))
            
        except Exception as e:
            logging.error(f"Ошибка парсинга PDF {file_path}: {e}")
            return None
    
    def extract_data_from_text(self, text, filename):
        """Извлечение данных из текста представления"""
        data = {
            'filename': filename,
            'takeoff_points': [],
            'landing_points': [],
            'flight_areas': [],
            'flight_info': {}
        }
        
        # Поиск точек взлета/посадки
        takeoff_landing_pattern = r'ВЗЛЕТ/ПОСАДКА\s+([\d\.]+[NS][\d\.]+[EW])\s+([\d\.]+[NS][\d\.]+[EW])'
        takeoff_matches = re.findall(takeoff_landing_pattern, text, re.IGNORECASE)
        
        for match in takeoff_matches:
            if len(match) >= 2:
                takeoff_coord = self.parse_coordinate(match[0])
                landing_coord = self.parse_coordinate(match[1])
                if takeoff_coord:
                    data['takeoff_points'].append(takeoff_coord)
                if landing_coord:
                    data['landing_points'].append(landing_coord)
        
        # Поиск координат окружностей (зоны полетов)
        circle_pattern = r'ОКРУЖНОСТЬ РАДИУС\s+(\d+)\s+КМ ЦЕНТР\s+([\d\.]+[NS][\d\.]+[EW])'
        circle_matches = re.findall(circle_pattern, text, re.IGNORECASE)
        
        for radius, center in circle_matches:
            center_coord = self.parse_coordinate(center)
            if center_coord:
                data['flight_areas'].append({
                    'type': 'circle',
                    'center': center_coord,
                    'radius_km': int(radius)
                })
        
        # Поиск полигонов (районов полетов)
        polygon_pattern = r'РАЙОН\s+((?:[\d\.]+[NS][\d\.]+[EW]\s*)+)'
        polygon_matches = re.findall(polygon_pattern, text, re.IGNORECASE)
        
        for polygon_coords in polygon_matches:
            coords_list = re.findall(r'[\d\.]+[NS][\d\.]+[EW]', polygon_coords)
            polygon_points = []
            for coord in coords_list:
                parsed_coord = self.parse_coordinate(coord)
                if parsed_coord:
                    polygon_points.append(parsed_coord)
            
            if len(polygon_points) >= 3:
                data['flight_areas'].append({
                    'type': 'polygon',
                    'points': polygon_points
                })
        
        # Извлечение общей информации о полетах
        date_pattern = r'(\d{2}/\d{2}/\d{4})'
        date_matches = re.findall(date_pattern, text)
        if date_matches:
            data['flight_info']['dates'] = date_matches
        
        time_pattern = r'(\d{2}:\d{2})\s*–\s*(\d{2}:\d{2})'
        time_matches = re.findall(time_pattern, text)
        if time_matches:
            data['flight_info']['flight_times'] = time_matches
        
        return data
    
    def parse_coordinate(self, coord_str):
        """Парсинг координат из строкового формата"""
        try:
            # Формат: 564144N0523226E
            lat_match = re.search(r'(\d{2})(\d{2})(\d{2})([NS])', coord_str)
            lon_match = re.search(r'(\d{3})(\d{2})(\d{2})([EW])', coord_str)
            
            if lat_match and lon_match:
                lat_deg = int(lat_match.group(1))
                lat_min = int(lat_match.group(2))
                lat_sec = int(lat_match.group(3))
                lat_dir = lat_match.group(4)
                
                lon_deg = int(lon_match.group(1))
                lon_min = int(lon_match.group(2))
                lon_sec = int(lon_match.group(3))
                lon_dir = lon_match.group(4)
                
                # Преобразование в десятичные градусы
                lat_decimal = lat_deg + lat_min/60 + lat_sec/3600
                lon_decimal = lon_deg + lon_min/60 + lon_sec/3600
                
                if lat_dir == 'S':
                    lat_decimal = -lat_decimal
                if lon_dir == 'W':
                    lon_decimal = -lon_decimal
                
                return {
                    'original': coord_str,
                    'decimal': (lat_decimal, lon_decimal),
                    'degrees_minutes_seconds': {
                        'lat': f"{lat_deg}°{lat_min:02d}'{lat_sec:02d}\"{lat_dir}",
                        'lon': f"{lon_deg}°{lon_min:02d}'{lon_sec:02d}\"{lon_dir}"
                    },
                    'degrees_minutes': {
                        'lat': f"{lat_deg}°{lat_min:02d}.{int(lat_sec/60*100):02d}'{lat_dir}",
                        'lon': f"{lon_deg}°{lon_min:02d}.{int(lon_sec/60*100):02d}'{lon_dir}"
                    }
                }
        except Exception as e:
            logging.error(f"Ошибка парсинга координаты {coord_str}: {e}")
        
        return None
    
    def create_kml_data(self, all_data):
        """Создание KML данных с правильным порядком координат"""
        kml = simplekml.Kml()
        
        for data in all_data:
            filename = data['filename']
            
            # Точки взлета - создаем как круговые полигоны
            for i, point in enumerate(data['takeoff_points']):
                # Создаем круговой полигон для точки взлета
                circle_points = self.create_circle_points(
                    point['decimal'][0],  # lat
                    point['decimal'][1],  # lon
                    self.takeoff_landing_radius.get()  # радиус из настроек
                )
                pol = kml.newpolygon(
                    name=f"🛫 Взлет {i+1} - {filename}",
                    outerboundaryis=circle_points
                )
                pol.style.polystyle.color = simplekml.Color.changealphaint(80, simplekml.Color.green)
                pol.style.linestyle.color = simplekml.Color.green
                pol.style.linestyle.width = 3
                pol.description = (
                    f"Точка взлета {i+1}\n"
                    f"Файл: {filename}\n"
                    f"Координаты: {point['original']}\n"
                    f"Широта: {point['decimal'][0]:.6f}\n"
                    f"Долгота: {point['decimal'][1]:.6f}\n"
                    f"Радиус: {self.takeoff_landing_radius.get()} км"
                )
            
            # Точки посадки - создаем как круговые полигоны
            for i, point in enumerate(data['landing_points']):
                # Создаем круговой полигон для точки посадки
                circle_points = self.create_circle_points(
                    point['decimal'][0],  # lat
                    point['decimal'][1],  # lon
                    self.takeoff_landing_radius.get()  # радиус из настроек
                )
                pol = kml.newpolygon(
                    name=f"🛬 Посадка {i+1} - {filename}",
                    outerboundaryis=circle_points
                )
                pol.style.polystyle.color = simplekml.Color.changealphaint(80, simplekml.Color.red)
                pol.style.linestyle.color = simplekml.Color.red
                pol.style.linestyle.width = 3
                pol.description = (
                    f"Точка посадки {i+1}\n"
                    f"Файл: {filename}\n"
                    f"Координаты: {point['original']}\n"
                    f"Широта: {point['decimal'][0]:.6f}\n"
                    f"Долгота: {point['decimal'][1]:.6f}\n"
                    f"Радиус: {self.takeoff_landing_radius.get()} км"
                )
            
            # Зоны полетов
            for i, area in enumerate(data['flight_areas']):
                if area['type'] == 'circle':
                    # Для кругов создаем полигон с правильной геометрией
                    circle_points = self.create_circle_points(
                        area['center']['decimal'][0],  # lat
                        area['center']['decimal'][1],  # lon
                        area['radius_km']  # радиус из представления
                    )
                    pol = kml.newpolygon(
                        name=f"🎯 Зона полетов {i+1} - {filename}",
                        outerboundaryis=circle_points
                    )
                    pol.style.polystyle.color = simplekml.Color.changealphaint(60, simplekml.Color.blue)
                    pol.style.linestyle.color = simplekml.Color.blue
                    pol.style.linestyle.width = 2
                    pol.description = (
                        f"Круговая зона полетов {i+1}\n"
                        f"Файл: {filename}\n"
                        f"Центр: {area['center']['original']}\n"
                        f"Радиус: {area['radius_km']} км"
                    )
                    
                elif area['type'] == 'polygon':
                    # ПРАВИЛЬНЫЙ порядок для KML полигонов: (longitude, latitude)
                    poly_coords = [(p['decimal'][1], p['decimal'][0]) for p in area['points']]
                    pol = kml.newpolygon(
                        name=f"📐 Район полетов {i+1} - {filename}",
                        outerboundaryis=poly_coords
                    )
                    pol.style.polystyle.color = simplekml.Color.changealphaint(80, simplekml.Color.yellow)
                    pol.style.linestyle.color = simplekml.Color.orange
                    pol.style.linestyle.width = 3
                    pol.description = (
                        f"Полигональная зона полетов {i+1}\n"
                        f"Файл: {filename}\n"
                        f"Точек в полигоне: {len(area['points'])}"
                    )
        
        return kml
    
    def create_circle_points(self, lat, lon, radius_km, points=36):
        """Создание точек для круговой зоны с правильной геометрией"""
        coords = []
        R = 6371.0  # Радиус Земли в км
        
        for i in range(points + 1):  # +1 для замыкания круга
            angle = 2.0 * math.pi * i / points
            
            # Вычисление новой точки с учетом сферической геометрии Земли
            lat_rad = math.radians(lat)
            lon_rad = math.radians(lon)
            
            # Формула для точки на заданном расстоянии от центра
            new_lat = math.asin(math.sin(lat_rad) * math.cos(radius_km/R) + 
                               math.cos(lat_rad) * math.sin(radius_km/R) * math.cos(angle))
            new_lon = lon_rad + math.atan2(math.sin(angle) * math.sin(radius_km/R) * math.cos(lat_rad),
                                         math.cos(radius_km/R) - math.sin(lat_rad) * math.sin(new_lat))
            
            # Преобразование обратно в градусы
            new_lat_deg = math.degrees(new_lat)
            new_lon_deg = math.degrees(new_lon)
            
            # ПРАВИЛЬНЫЙ порядок для KML: (longitude, latitude)
            coords.append((new_lon_deg, new_lat_deg))
        
        return coords
    
    def _processing_complete(self):
        """Завершение обработки"""
        self.update_status("Обработка завершена")
        messagebox.showinfo("Успех", "Файлы успешно обработаны. Можете экспортировать KML.")
    
    def export_kml(self):
        """Экспорт KML файла"""
        if not self.kml_data:
            messagebox.showwarning("Внимание", "Сначала обработайте файлы")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Сохранить KML файл",
            defaultextension=".kml",
            filetypes=[("KML files", "*.kml"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                # Добавляем метаданные для лучшей совместимости
                self.kml_data.document.name = "Точки взлета и посадки БВС"
                self.kml_data.document.description = "Сгенерировано EGOK Renamer PDF→KML плагином"
                
                # Сохраняем с правильной кодировкой
                self.kml_data.save(file_path)
                
                self.update_status(f"KML файл сохранен: {file_path}")
                
                # Показываем подсказку для SAS.Planet
                messagebox.showinfo(
                    "Успех", 
                    f"KML файл успешно сохранен:\n{file_path}\n\n"
                    f"Рекомендации для SAS.Planet:\n"
                    f"1. Откройте файл через меню 'Метки'\n"
                    f"2. Точки взлета/посадки отображаются как круги\n"
                    f"3. Радиус точек: {self.takeoff_landing_radius.get()} км"
                )
            except Exception as e:
                messagebox.showerror("Ошибка", f"Ошибка сохранения KML: {str(e)}")
    
    def show_coordinates(self):
        """Показать извлеченные координаты"""
        if not self.loaded_files:
            messagebox.showwarning("Внимание", "Нет загруженных файлов")
            return
        
        self.result_text.delete(1.0, tk.END)
        
        for file_path in self.loaded_files:
            try:
                with open(file_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    text = ""
                    
                    for page in pdf_reader.pages:
                        text += page.extract_text()
                
                self.result_text.insert(tk.END, f"\n=== {os.path.basename(file_path)} ===\n")
                
                # Поиск всех координат в тексте
                coord_pattern = r'[\d\.]+[NS][\d\.]+[EW]'
                coords = re.findall(coord_pattern, text)
                
                for coord in coords[:10]:  # Показываем первые 10 координат
                    parsed = self.parse_coordinate(coord)
                    if parsed:
                        format_type = self.coordinate_format.get()
                        if format_type == "degrees":
                            display_coord = f"{parsed['decimal'][0]:.6f}, {parsed['decimal'][1]:.6f}"
                        elif format_type == "degrees_minutes":
                            display_coord = f"{parsed['degrees_minutes']['lat']}, {parsed['degrees_minutes']['lon']}"
                        else:
                            display_coord = f"{parsed['degrees_minutes_seconds']['lat']}, {parsed['degrees_minutes_seconds']['lon']}"
                        
                        self.result_text.insert(tk.END, f"{coord} → {display_coord}\n")
                
                if len(coords) > 10:
                    self.result_text.insert(tk.END, f"... и еще {len(coords) - 10} координат\n")
                    
            except Exception as e:
                self.result_text.insert(tk.END, f"Ошибка чтения файла: {str(e)}\n")
    
    def update_status(self, message):
        """Обновление статусной строки"""
        self.status_var.set(message)
        self.root.update_idletasks()

def get_plugin_class():
    return PDFDecoderPlugin