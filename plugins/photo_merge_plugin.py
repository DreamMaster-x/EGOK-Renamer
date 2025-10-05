# plugins/photo_merge_plugin.py
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import shutil
from pathlib import Path
import threading
from datetime import datetime
from PIL import Image, ImageTk
import logging

class PhotoMergePlugin:
    """Плагин для объединения фотографий из нескольких папок с уникальным переименованием"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.source_folders = []
        self.is_processing = False
        
    def get_tab_name(self):
        return "Объединение фото"
    
    def create_tab(self):
        tab_frame = ttk.Frame(self.root)
        self.create_interface(tab_frame)
        return tab_frame
    
    def create_interface(self, parent):
        # Основной контейнер с прокруткой
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Объединение фотографий из нескольких папок", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Описание
        desc_text = """ПРОБЛЕМА: Фотоаппарат создает новые папки при превышении 4000 файлов, 
и имена файлов повторяются в разных папках.

РЕШЕНИЕ: Этот плагин объединяет фотографии из нескольких папок в одну, 
присваивая уникальные имена и сохраняя порядок по дате создания."""
        
        desc_label = ttk.Label(main_frame, text=desc_text, justify=tk.LEFT, wraplength=600)
        desc_label.pack(pady=(0, 20))
        
        # Фрейм для папок-источников
        source_frame = ttk.LabelFrame(main_frame, text="Папки-источники с фотографиями")
        source_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Список папок-источников
        self.source_listbox = tk.Listbox(source_frame, height=6, selectmode=tk.SINGLE)
        self.source_listbox.pack(fill=tk.X, padx=5, pady=5)
        
        # Кнопки управления папками-источниками
        source_buttons_frame = ttk.Frame(source_frame)
        source_buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(source_buttons_frame, text="Добавить папку", 
                  command=self.add_source_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(source_buttons_frame, text="Удалить выбранную", 
                  command=self.remove_source_folder).pack(side=tk.LEFT, padx=2)
        ttk.Button(source_buttons_frame, text="Очистить все", 
                  command=self.clear_source_folders).pack(side=tk.LEFT, padx=2)
        
        # Фрейм для папки назначения
        dest_frame = ttk.LabelFrame(main_frame, text="Папка назначения для объединенных фотографий")
        dest_frame.pack(fill=tk.X, pady=(0, 15))
        
        dest_input_frame = ttk.Frame(dest_frame)
        dest_input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.dest_folder_var = tk.StringVar()
        ttk.Entry(dest_input_frame, textvariable=self.dest_folder_var, 
                 state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(dest_input_frame, text="Обзор", 
                  command=self.browse_dest_folder).pack(side=tk.RIGHT)
        
        # Настройки переименования
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки переименования")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Шаблон имени
        template_frame = ttk.Frame(settings_frame)
        template_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(template_frame, text="Шаблон имени:").pack(side=tk.LEFT)
        
        self.template_var = tk.StringVar(value="photo_{date}_{counter:06d}")
        template_entry = ttk.Entry(template_frame, textvariable=self.template_var, width=30)
        template_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Опции
        options_frame = ttk.Frame(settings_frame)
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.preserve_structure_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Сохранить структуру подпапок", 
                       variable=self.preserve_structure_var).pack(side=tk.LEFT, padx=5)
        
        self.keep_original_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Сохранить оригинальные файлы (копировать)", 
                       variable=self.keep_original_var).pack(side=tk.LEFT, padx=5)
        
        # Расширения файлов
        ext_frame = ttk.Frame(settings_frame)
        ext_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(ext_frame, text="Расширения:").pack(side=tk.LEFT)
        
        self.extensions_var = tk.StringVar(value="jpg,jpeg,png,cr2,nef,arw,dng")
        ext_entry = ttk.Entry(ext_frame, textvariable=self.extensions_var, width=30)
        ext_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Прогресс
        progress_frame = ttk.LabelFrame(main_frame, text="Прогресс выполнения")
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        status_label.pack(padx=5, pady=(0, 5))
        
        # Лог
        log_frame = ttk.Frame(progress_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки выполнения
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.merge_button = ttk.Button(button_frame, text="Начать объединение", 
                                      command=self.start_merge_process)
        self.merge_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Очистить лог", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=5)
        
        # Подсказка
        help_text = """ПОДСКАЗКА:
• Добавьте все папки с фотографиями (включая вложенные папки если нужно)
• Выберите папку для сохранения результатов
• Файлы будут отсортированы по дате создания и переименованы
• Шаблон имени: {date} - дата создания, {counter} - порядковый номер"""
        
        help_label = ttk.Label(main_frame, text=help_text, justify=tk.LEFT, 
                              foreground="gray", font=('Arial', 8))
        help_label.pack(pady=10)
    
    def add_source_folder(self):
        """Добавить папку-источник"""
        folder = filedialog.askdirectory(title="Выберите папку с фотографиями")
        if folder and folder not in self.source_folders:
            self.source_folders.append(folder)
            self.source_listbox.insert(tk.END, folder)
            self.log_message(f"Добавлена папка: {folder}")
    
    def remove_source_folder(self):
        """Удалить выбранную папку-источник"""
        selection = self.source_listbox.curselection()
        if selection:
            index = selection[0]
            folder = self.source_folders.pop(index)
            self.source_listbox.delete(index)
            self.log_message(f"Удалена папка: {folder}")
    
    def clear_source_folders(self):
        """Очистить все папки-источники"""
        self.source_folders.clear()
        self.source_listbox.delete(0, tk.END)
        self.log_message("Все папки-источники очищены")
    
    def browse_dest_folder(self):
        """Выбрать папку назначения"""
        folder = filedialog.askdirectory(title="Выберите папку для сохранения объединенных фотографий")
        if folder:
            self.dest_folder_var.set(folder)
    
    def log_message(self, message):
        """Добавить сообщение в лог"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_var.set(message)
    
    def clear_log(self):
        """Очистить лог"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_var.set("Лог очищен")
    
    def start_merge_process(self):
        """Начать процесс объединения"""
        if self.is_processing:
            messagebox.showwarning("Внимание", "Процесс уже выполняется!")
            return
        
        if not self.source_folders:
            messagebox.showwarning("Внимание", "Добавьте хотя бы одну папку-источник!")
            return
        
        if not self.dest_folder_var.get():
            messagebox.showwarning("Внимание", "Выберите папку назначения!")
            return
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.merge_photos)
        thread.daemon = True
        thread.start()
    
    def merge_photos(self):
        """Основная логика объединения фотографий"""
        self.is_processing = True
        self.merge_button.config(state='disabled')
        
        try:
            dest_folder = self.dest_folder_var.get()
            extensions = [ext.strip().lower() for ext in self.extensions_var.get().split(",")]
            
            # Создаем папку назначения если не существует
            os.makedirs(dest_folder, exist_ok=True)
            
            self.log_message("Начало сканирования папок...")
            
            # Собираем все файлы из всех папок-источников
            all_files = []
            for source_folder in self.source_folders:
                if self.preserve_structure_var.get():
                    # Рекурсивный обход с сохранением структуры
                    for root, dirs, files in os.walk(source_folder):
                        for file in files:
                            file_ext = Path(file).suffix.lower().lstrip('.')
                            if file_ext in extensions:
                                file_path = os.path.join(root, file)
                                all_files.append(file_path)
                else:
                    # Только файлы в корне папок-источников
                    if os.path.exists(source_folder):
                        for file in os.listdir(source_folder):
                            file_path = os.path.join(source_folder, file)
                            if os.path.isfile(file_path):
                                file_ext = Path(file).suffix.lower().lstrip('.')
                                if file_ext in extensions:
                                    all_files.append(file_path)
            
            if not all_files:
                self.log_message("Фотографии не найдены!")
                return
            
            self.log_message(f"Найдено {len(all_files)} файлов")
            
            # Сортируем файлы по дате создания
            files_with_dates = []
            for file_path in all_files:
                try:
                    create_time = os.path.getctime(file_path)
                    files_with_dates.append((file_path, create_time))
                except Exception as e:
                    self.log_message(f"Ошибка получения даты для {file_path}: {e}")
                    files_with_dates.append((file_path, 0))
            
            # Сортируем по дате создания
            files_with_dates.sort(key=lambda x: x[1])
            
            self.log_message("Начало обработки файлов...")
            
            # Обрабатываем файлы
            processed_count = 0
            total_files = len(files_with_dates)
            
            for i, (file_path, create_time) in enumerate(files_with_dates):
                try:
                    # Обновляем прогресс
                    progress = (i + 1) / total_files * 100
                    self.root.after(0, lambda: self.progress_var.set(progress))
                    
                    original_name = Path(file_path).name
                    file_ext = Path(file_path).suffix.lower().lstrip('.')
                    
                    # Генерируем новое имя
                    date_str = datetime.fromtimestamp(create_time).strftime("%Y%m%d_%H%M%S")
                    new_name = self.template_var.get().format(
                        date=date_str,
                        counter=i+1,
                        extension=file_ext
                    ) + f".{file_ext}"
                    
                    new_path = os.path.join(dest_folder, new_name)
                    
                    # Если файл с таким именем уже существует, добавляем суффикс
                    counter = 1
                    while os.path.exists(new_path):
                        name_part = new_name.rsplit('.', 1)[0]
                        new_name = f"{name_part}_{counter:02d}.{file_ext}"
                        new_path = os.path.join(dest_folder, new_name)
                        counter += 1
                    
                    # Копируем или перемещаем файл
                    if self.keep_original_var.get():
                        shutil.copy2(file_path, new_path)
                        action = "Скопирован"
                    else:
                        shutil.move(file_path, new_path)
                        action = "Перемещен"
                    
                    self.log_message(f"{action}: {original_name} -> {new_name}")
                    processed_count += 1
                    
                except Exception as e:
                    self.log_message(f"Ошибка обработки {file_path}: {e}")
            
            # Завершение
            self.root.after(0, lambda: self.progress_var.set(100))
            self.log_message(f"Обработка завершена! Обработано {processed_count} из {total_files} файлов")
            
            if processed_count == total_files:
                self.log_message("✓ Все файлы успешно обработаны!")
            else:
                self.log_message(f"⚠ Обработано {processed_count} из {total_files} файлов")
                
        except Exception as e:
            self.log_message(f"Критическая ошибка: {e}")
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.merge_button.config(state='normal'))

# Обязательная функция для загрузки плагина
def get_plugin_class():
    return PhotoMergePlugin