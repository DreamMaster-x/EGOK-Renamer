# plugins/file_generator_plugin.py
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import shutil
from pathlib import Path
import threading
from datetime import datetime, timedelta
import random
import string

# Проверяем наличие PIL
try:
    from PIL import Image, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

class FileGeneratorPlugin:
    """Плагин для генерации тестовых файлов"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.is_processing = False
        
    def get_tab_name(self):
        return "Генератор файлов"
    
    def create_tab(self):
        tab_frame = ttk.Frame(self.root)
        self.create_interface(tab_frame)
        return tab_frame
    
    def create_interface(self, parent):
        # Основной контейнер с прокруткой
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Генератор тестовых файлов", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 15))
        
        # Описание
        desc_text = "Создает тестовые файлы для проверки работы других плагинов."
        desc_label = ttk.Label(main_frame, text=desc_text, justify=tk.LEFT, wraplength=600)
        desc_label.pack(pady=(0, 20))
        
        # Настройки генерации
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки генерации")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Количество файлов
        count_frame = ttk.Frame(settings_frame)
        count_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(count_frame, text="Количество файлов:").pack(side=tk.LEFT)
        
        self.file_count_var = tk.StringVar(value="10")
        count_entry = ttk.Entry(count_frame, textvariable=self.file_count_var, width=10)
        count_entry.pack(side=tk.LEFT, padx=5)
        
        # Тип файлов
        type_frame = ttk.Frame(settings_frame)
        type_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(type_frame, text="Тип файлов:").pack(side=tk.LEFT)
        
        self.file_type_var = tk.StringVar(value="text")
        ttk.Radiobutton(type_frame, text="Текстовые", 
                       variable=self.file_type_var, value="text").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(type_frame, text="Изображения", 
                       variable=self.file_type_var, value="images").pack(side=tk.LEFT, padx=5)
        
        # Папка назначения
        dest_frame = ttk.Frame(settings_frame)
        dest_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(dest_frame, text="Папка назначения:").pack(side=tk.LEFT)
        
        self.dest_folder_var = tk.StringVar()
        ttk.Entry(dest_frame, textvariable=self.dest_folder_var, 
                 state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(dest_frame, text="Обзор", 
                  command=self.browse_dest_folder).pack(side=tk.RIGHT)
        
        # Шаблон имени
        template_frame = ttk.Frame(settings_frame)
        template_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(template_frame, text="Шаблон имени:").pack(side=tk.LEFT)
        
        self.template_var = tk.StringVar(value="test_file_{counter:04d}")
        template_entry = ttk.Entry(template_frame, textvariable=self.template_var, width=30)
        template_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
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
        
        self.log_text = tk.Text(log_frame, height=6, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Кнопки выполнения
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.generate_button = ttk.Button(button_frame, text="Начать генерацию", 
                                        command=self.start_generation_process)
        self.generate_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Очистить лог", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=5)
        
        # Информация о PIL
        if not PIL_AVAILABLE:
            warning_text = "Внимание: PIL не установлен. Генерация изображений недоступна."
            warning_label = ttk.Label(main_frame, text=warning_text, 
                                     foreground="red", font=('Arial', 8))
            warning_label.pack(pady=5)
    
    def browse_dest_folder(self):
        """Выбрать папку назначения"""
        folder = filedialog.askdirectory(title="Выберите папку для сохранения сгенерированных файлов")
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
    
    def generate_random_string(self, length=8):
        """Генерация случайной строки"""
        return ''.join(random.choices(string.ascii_letters + string.digits, k=length))
    
    def create_test_image(self, file_path, counter):
        """Создание тестового изображения"""
        if not PIL_AVAILABLE:
            self.log_message("Ошибка: PIL не установлен, невозможно создать изображение")
            return False
            
        try:
            # Создаем простое изображение
            width, height = 800, 600
            image = Image.new('RGB', (width, height), color=(
                random.randint(0, 255), 
                random.randint(0, 255), 
                random.randint(0, 255)
            ))
            
            draw = ImageDraw.Draw(image)
            # Простой текст
            text = f"Test Image {counter}"
            draw.text((50, 50), text, fill=(255, 255, 255))
            
            image.save(file_path, 'JPEG', quality=95)
            return True
        except Exception as e:
            self.log_message(f"Ошибка создания изображения: {e}")
            return False
    
    def create_test_text_file(self, file_path, counter):
        """Создание тестового текстового файла"""
        try:
            content = f"Тестовый файл №{counter}\nСгенерирован: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nСлучайный ID: {self.generate_random_string()}"
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            return True
        except Exception as e:
            self.log_message(f"Ошибка создания текстового файла: {e}")
            return False
    
    def start_generation_process(self):
        """Начать процесс генерации"""
        if self.is_processing:
            messagebox.showwarning("Внимание", "Процесс уже выполняется!")
            return
        
        # Проверка входных данных
        try:
            file_count = int(self.file_count_var.get())
            if file_count <= 0:
                messagebox.showwarning("Внимание", "Количество файлов должно быть положительным числом!")
                return
        except ValueError:
            messagebox.showwarning("Внимание", "Введите корректное число файлов!")
            return
        
        if not self.dest_folder_var.get():
            messagebox.showwarning("Внимание", "Выберите папку назначения!")
            return
        
        if self.file_type_var.get() == "images" and not PIL_AVAILABLE:
            messagebox.showwarning("Внимание", "PIL не установлен. Генерация изображений недоступна!")
            return
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.generate_files)
        thread.daemon = True
        thread.start()
    
    def generate_files(self):
        """Основная логика генерации файлов"""
        self.is_processing = True
        self.generate_button.config(state='disabled')
        
        try:
            dest_folder = self.dest_folder_var.get()
            file_count = int(self.file_count_var.get())
            file_type = self.file_type_var.get()
            
            # Создаем папку назначения если не существует
            os.makedirs(dest_folder, exist_ok=True)
            
            self.log_message(f"Начало генерации {file_count} {file_type} файлов...")
            
            successful_count = 0
            
            for i in range(1, file_count + 1):
                try:
                    # Обновляем прогресс
                    progress = (i) / file_count * 100
                    self.root.after(0, lambda p=progress: self.progress_var.set(p))
                    
                    # Генерируем имя файла
                    file_name = self.template_var.get().format(counter=i) 
                    
                    if file_type == "images":
                        file_name += ".jpg"
                    else:
                        file_name += ".txt"
                    
                    file_path = os.path.join(dest_folder, file_name)
                    
                    # Создаем файл в зависимости от типа
                    if file_type == "images":
                        success = self.create_test_image(file_path, i)
                    else:
                        success = self.create_test_text_file(file_path, i)
                    
                    if success:
                        successful_count += 1
                        self.log_message(f"Создан: {file_name}")
                    else:
                        self.log_message(f"Ошибка создания: {file_name}")
                    
                except Exception as e:
                    self.log_message(f"Ошибка при создании файла #{i}: {e}")
            
            # Завершение
            self.root.after(0, lambda: self.progress_var.set(100))
            self.log_message(f"Генерация завершена! Успешно создано {successful_count} из {file_count} файлов")
            
        except Exception as e:
            self.log_message(f"Критическая ошибка: {e}")
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.generate_button.config(state='normal'))

# Обязательная функция для загрузки плагина
def get_plugin_class():
    return FileGeneratorPlugin