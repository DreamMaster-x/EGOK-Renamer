# plugins/telemetry_plugin.py
import os
import json
import zipfile
import threading
import queue
import logging
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS

class TelemetryPlugin:
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.log_queue = queue.Queue()
        self.setup_plugin_settings()
    
    def setup_plugin_settings(self):
        """Инициализация настроек плагина"""
        plugin_settings = self.settings.settings.get("telemetry_plugin", {})
        
        # Настройки по умолчанию для плагина
        default_plugin_settings = {
            "telemetry_folder": "",
            "photos_folder": "",
            "output_telemetry_name": "tele_photo.tlm",
            "archive_template": "Дата_{номер маршрута}",
            "compress_to_zip": True,
            "telemetry_folder_history": [],
            "photos_folder_history": [],
            "output_name_history": ["tele_photo.tlm"],
            "archive_template_history": ["Дата_{номер маршрута}"],
            "route_number": "M2.1"
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
        """Создание вкладки плагина"""
        tab_frame = ttk.Frame(self.root)
        
        # Основной фрейм с прокруткой
        main_frame = ttk.Frame(tab_frame)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Заголовок
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
        
        telemetry_var = tk.StringVar(value=self.plugin_settings["telemetry_folder"])
        self.telemetry_var = telemetry_var
        
        telemetry_combo = ttk.Combobox(
            telemetry_frame, 
            textvariable=telemetry_var,
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
        
        photos_var = tk.StringVar(value=self.plugin_settings["photos_folder"])
        self.photos_var = photos_var
        
        photos_combo = ttk.Combobox(
            photos_frame, 
            textvariable=photos_var,
            values=self.plugin_settings["photos_folder_history"],
            width=50
        )
        photos_combo.pack(fill=tk.X, pady=2)
        ttk.Button(photos_frame, text="Обзор", 
                  command=self.browse_photos_folder).pack(anchor=tk.W, pady=2)
        
        # Настройки выходных файлов
        output_frame = ttk.Frame(settings_frame)
        output_frame.pack(fill=tk.X, pady=5, padx=5)
        
        # Имя выходного файла телеметрии
        ttk.Label(output_frame, text="Имя выходного файла телеметрии:").pack(anchor=tk.W)
        
        output_name_var = tk.StringVar(value=self.plugin_settings["output_telemetry_name"])
        self.output_name_var = output_name_var
        
        output_name_combo = ttk.Combobox(
            output_frame, 
            textvariable=output_name_var,
            values=self.plugin_settings["output_name_history"],
            width=30
        )
        output_name_combo.pack(fill=tk.X, pady=2)
        
        # Шаблон архива
        ttk.Label(output_frame, text="Шаблон имени архива:").pack(anchor=tk.W, pady=(10, 0))
        
        archive_var = tk.StringVar(value=self.plugin_settings["archive_template"])
        self.archive_var = archive_var
        
        archive_combo = ttk.Combobox(
            output_frame, 
            textvariable=archive_var,
            values=self.plugin_settings["archive_template_history"],
            width=30
        )
        archive_combo.pack(fill=tk.X, pady=2)
        
        # Номер маршрута
        route_frame = ttk.Frame(output_frame)
        route_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(route_frame, text="Номер маршрута:").pack(side=tk.LEFT)
        
        route_var = tk.StringVar(value=self.plugin_settings["route_number"])
        self.route_var = route_var
        
        route_entry = ttk.Entry(route_frame, textvariable=route_var, width=15)
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
        
        # Запуск обработки очереди логов
        self.process_log_queue()
        
        return tab_frame
    
    def browse_telemetry_file(self):
        """Выбор файла телеметрии"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл телеметрии",
            filetypes=[("TLM files", "*.tlm"), ("All files", "*.*")]
        )
        if file_path:
            self.telemetry_var.set(file_path)
            self.add_to_history("telemetry_folder_history", file_path)
    
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
    
    def browse_photos_folder(self):
        """Выбор папки с фотографиями"""
        folder = filedialog.askdirectory(title="Выберите папку с фотографиями")
        if folder:
            self.photos_var.set(folder)
            self.add_to_history("photos_folder_history", folder)
    
    def add_to_history(self, history_key, value):
        """Добавление значения в историю"""
        if value and value not in self.plugin_settings[history_key]:
            self.plugin_settings[history_key].insert(0, value)
            # Ограничиваем историю 10 элементами
            self.plugin_settings[history_key] = self.plugin_settings[history_key][:10]
            self.save_plugin_settings()
    
    def save_settings(self):
        """Сохранение настроек плагина"""
        try:
            # Сохраняем текущие значения
            self.plugin_settings["telemetry_folder"] = self.telemetry_var.get()
            self.plugin_settings["photos_folder"] = self.photos_var.get()
            self.plugin_settings["output_telemetry_name"] = self.output_name_var.get()
            self.plugin_settings["archive_template"] = self.archive_var.get()
            self.plugin_settings["compress_to_zip"] = self.compress_var.get()
            self.plugin_settings["route_number"] = self.route_var.get()
            
            # Добавляем в историю
            self.add_to_history("output_name_history", self.output_name_var.get())
            self.add_to_history("archive_template_history", self.archive_var.get())
            
            self.save_plugin_settings()
            self.log_message("Настройки сохранены успешно")
            messagebox.showinfo("Успех", "Настройки плагина сохранены!")
            
        except Exception as e:
            self.log_message(f"Ошибка сохранения настроек: {e}", "error")
            messagebox.showerror("Ошибка", f"Ошибка сохранения настроек: {e}")
    
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
                            return dt.strftime('%Y/%m/%d'), dt.strftime('%H:%M:%S')
        except Exception as e:
            self.log_message(f"Ошибка чтения EXIF {image_path}: {e}", "warning")
        
        # Если EXIF нет, используем время создания файла
        try:
            file_time = datetime.fromtimestamp(os.path.getctime(image_path))
            return file_time.strftime('%Y/%m/%d'), file_time.strftime('%H:%M:%S')
        except:
            return "0000/00/00", "00:00:00"
    
    def parse_telemetry_file(self, file_path):
        """Парсинг файла телеметрии"""
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
                                telemetry_data.append({
                                    'datetime': dt,
                                    'line': line,
                                    'line_num': line_num
                                })
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
        """Поток обработки телеметрии"""
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
                        date_str, time_str = self.get_exif_datetime(photo_file)
                        
                        # Преобразуем в datetime для сравнения
                        try:
                            photo_datetime = datetime.strptime(f"{date_str} {time_str}", '%Y/%m/%d %H:%M:%S')
                        except:
                            # Если не удалось распарсить, используем текущее время
                            photo_datetime = datetime.now()
                            self.log_message(f"Использовано текущее время для {photo_file.name}", "warning")
                        
                        # Ищем ближайшую запись телеметрии
                        closest_telemetry = self.find_closest_telemetry(photo_datetime, telemetry_data)
                        
                        if closest_telemetry:
                            # Формируем строку для выходного файла
                            # Формат: имя_файла дата время данные_телеметрии
                            telemetry_parts = closest_telemetry['line'].split()[3:]  # Пропускаем 'L', дату, время
                            telemetry_str = ' '.join(telemetry_parts[:10])  # Берем первые 10 значений
                            
                            output_line = f"{photo_file.name}\t{date_str}\t{time_str}\t{telemetry_str}\n"
                            out_file.write(output_line)
                            processed_count += 1
                            
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
                # Добавляем все фотографии
                photo_extensions = ['.jpg', '.jpeg', '.png', '.JPG', '.JPEG', '.PNG', '.tlm']
                for ext in photo_extensions:
                    for file_path in Path(photos_folder).glob(f"*{ext}"):
                        if file_path.name != archive_name:  # Не добавляем сам архив
                            zipf.write(file_path, file_path.name)
            
            self.log_message(f"Создан архив: {archive_name} ({file_count} файлов)")
            
        except Exception as e:
            self.log_message(f"Ошибка создания архива: {e}", "error")
    
    def log_message(self, message, level="info"):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        
        # Добавляем в очередь для безопасного обновления GUI
        self.log_queue.put((log_entry, level))
    
    def process_log_queue(self):
        """Обработка очереди логов"""
        try:
            while True:
                log_entry, level = self.log_queue.get_nowait()
                
                self.log_text.config(state=tk.NORMAL)
                
                # Определяем цвет в зависимости от уровня
                if level == "error":
                    self.log_text.insert(tk.END, log_entry, "error")
                elif level == "warning":
                    self.log_text.insert(tk.END, log_entry, "warning")
                else:
                    self.log_text.insert(tk.END, log_entry, "info")
                
                # Автопрокрутка
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
                
        except queue.Empty:
            pass
        finally:
            # Планируем следующую проверку
            self.root.after(100, self.process_log_queue)
    
    def clear_logs(self):
        """Очистка логов"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

def get_plugin_class():
    return TelemetryPlugin