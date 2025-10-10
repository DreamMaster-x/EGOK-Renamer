# plugins/comma_replacer_plugin.py
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from pathlib import Path

class CommaReplacerPlugin:
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.is_processing = False
        self.stop_requested = False
        
    def get_tab_name(self):
        return "Замена запятых"
    
    def create_tab(self):
        tab_frame = ttk.Frame(self.root)
        
        # Основной заголовок
        title_label = ttk.Label(tab_frame, text="Замена запятых на точки в текстовых файлах", 
                               font=('Arial', 12, 'bold'))
        title_label.pack(pady=10)
        
        # Фрейм настроек
        settings_frame = ttk.LabelFrame(tab_frame, text="Настройки обработки")
        settings_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Выбор папки
        folder_frame = ttk.Frame(settings_frame)
        folder_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(folder_frame, text="Папка с файлами:").pack(side=tk.LEFT)
        
        self.folder_var = tk.StringVar()
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=50)
        folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Button(folder_frame, text="Обзор", command=self.browse_folder).pack(side=tk.LEFT)
        
        # Расширения файлов
        ext_frame = ttk.Frame(settings_frame)
        ext_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(ext_frame, text="Расширения:").pack(side=tk.LEFT)
        
        self.ext_var = tk.StringVar(value="txt,csv,log,ini,cfg,xml,json")
        ext_entry = ttk.Entry(ext_frame, textvariable=self.ext_var, width=50)
        ext_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Опции обработки
        options_frame = ttk.Frame(settings_frame)
        options_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Создавать резервные копии", 
                       variable=self.backup_var).pack(side=tk.LEFT, padx=5)
        
        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Обрабатывать вложенные папки", 
                       variable=self.recursive_var).pack(side=tk.LEFT, padx=5)
        
        # Кнопки управления
        button_frame = ttk.Frame(tab_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.process_btn = ttk.Button(button_frame, text="Начать обработку", 
                                     command=self.start_processing)
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(button_frame, text="Остановить", 
                                  command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Очистить лог", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(tab_frame, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        
        # Статус
        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = ttk.Label(tab_frame, textvariable=self.status_var)
        status_label.pack(pady=5)
        
        # Лог обработки
        log_frame = ttk.LabelFrame(tab_frame, text="Лог обработки")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Текстовое поле для лога с прокруткой
        self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=15)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Статистика
        stats_frame = ttk.Frame(tab_frame)
        stats_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.stats_var = tk.StringVar(value="Обработано файлов: 0 | Замен: 0")
        stats_label = ttk.Label(stats_frame, textvariable=self.stats_var)
        stats_label.pack()
        
        return tab_frame
    
    def browse_folder(self):
        """Выбор папки для обработки"""
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)
    
    def start_processing(self):
        """Запуск обработки файлов в отдельном потоке"""
        folder = self.folder_var.get()
        if not folder or not os.path.exists(folder):
            messagebox.showerror("Ошибка", "Укажите существующую папку для обработки")
            return
        
        self.is_processing = True
        self.stop_requested = False
        self.process_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress['value'] = 0
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def stop_processing(self):
        """Остановка обработки"""
        self.stop_requested = True
        self.status_var.set("Остановка...")
    
    def clear_log(self):
        """Очистка лога"""
        self.log_text.delete(1.0, tk.END)
    
    def log_message(self, message):
        """Добавление сообщения в лог"""
        self.root.after(0, lambda: self._add_log_message(message))
    
    def _add_log_message(self, message):
        """Добавление сообщения в лог (вызывается из главного потока)"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
    
    def update_progress(self, value, max_value):
        """Обновление прогресса"""
        self.root.after(0, lambda: self._update_progress_value(value, max_value))
    
    def _update_progress_value(self, value, max_value):
        """Обновление прогресса (вызывается из главного потока)"""
        if max_value > 0:
            self.progress['value'] = (value / max_value) * 100
    
    def update_status(self, status):
        """Обновление статуса"""
        self.root.after(0, lambda: self.status_var.set(status))
    
    def update_stats(self, files_processed, replacements_count):
        """Обновление статистики"""
        self.root.after(0, lambda: self._update_stats_value(files_processed, replacements_count))
    
    def _update_stats_value(self, files_processed, replacements_count):
        """Обновление статистики (вызывается из главного потока)"""
        self.stats_var.set(f"Обработано файлов: {files_processed} | Замен: {replacements_count}")
    
    def processing_finished(self):
        """Завершение обработки"""
        self.root.after(0, self._finish_processing)
    
    def _finish_processing(self):
        """Завершение обработки (вызывается из главного потока)"""
        self.is_processing = False
        self.process_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress['value'] = 100
        
        if self.stop_requested:
            self.status_var.set("Обработка остановлена")
        else:
            self.status_var.set("Обработка завершена")
    
    def process_files(self):
        """Основная функция обработки файлов"""
        try:
            folder = self.folder_var.get()
            extensions = [ext.strip().lower() for ext in self.ext_var.get().split(",")]
            create_backup = self.backup_var.get()
            recursive = self.recursive_var.get()
            
            # Собираем список файлов для обработки
            files_to_process = []
            if recursive:
                for root_dir, _, files in os.walk(folder):
                    for file in files:
                        file_ext = Path(file).suffix.lower()[1:]
                        if file_ext in extensions:
                            files_to_process.append(os.path.join(root_dir, file))
            else:
                for file in os.listdir(folder):
                    file_path = os.path.join(folder, file)
                    if os.path.isfile(file_path):
                        file_ext = Path(file).suffix.lower()[1:]
                        if file_ext in extensions:
                            files_to_process.append(file_path)
            
            total_files = len(files_to_process)
            files_processed = 0
            total_replacements = 0
            
            self.update_status(f"Найдено файлов для обработки: {total_files}")
            self.log_message(f"Начало обработки {total_files} файлов...")
            
            for file_path in files_to_process:
                if self.stop_requested:
                    break
                
                self.update_status(f"Обработка: {os.path.basename(file_path)}")
                self.update_progress(files_processed, total_files)
                
                try:
                    replacements = self.process_single_file(file_path, create_backup)
                    total_replacements += replacements
                    files_processed += 1
                    
                    if replacements > 0:
                        self.log_message(f"✓ {os.path.basename(file_path)} - замен: {replacements}")
                    else:
                        self.log_message(f"- {os.path.basename(file_path)} - замен нет")
                    
                    self.update_stats(files_processed, total_replacements)
                    
                except Exception as e:
                    self.log_message(f"✗ Ошибка в {os.path.basename(file_path)}: {str(e)}")
            
            # Финальное обновление
            self.update_progress(files_processed, total_files)
            self.update_stats(files_processed, total_replacements)
            
            if self.stop_requested:
                self.log_message("Обработка остановлена пользователем")
            else:
                self.log_message(f"Обработка завершена. Всего замен: {total_replacements}")
            
            self.processing_finished()
            
        except Exception as e:
            self.log_message(f"Критическая ошибка: {str(e)}")
            self.processing_finished()
    
    def process_single_file(self, file_path, create_backup):
        """Обработка одного файла"""
        replacements_count = 0
        
        # Создаем резервную копию если нужно
        backup_path = None
        if create_backup:
            backup_path = file_path + ".backup"
            import shutil
            shutil.copy2(file_path, backup_path)
        
        try:
            # Читаем файл с определением кодировки
            content, encoding = self.read_file_with_encoding(file_path)
            
            if content is None:
                raise Exception("Не удалось прочитать файл")
            
            # Заменяем запятые на точки
            new_content, replacements = self.replace_commas_with_dots(content)
            replacements_count = replacements
            
            if replacements > 0:
                # Сохраняем изменения
                self.write_file_with_encoding(file_path, new_content, encoding)
            
            # Удаляем резервную копию если изменений не было
            if create_backup and replacements_count == 0 and backup_path and os.path.exists(backup_path):
                os.remove(backup_path)
                
        except Exception as e:
            # Восстанавливаем из резервной копии при ошибке
            if create_backup and backup_path and os.path.exists(backup_path):
                import shutil
                shutil.move(backup_path, file_path)
            raise e
        
        return replacements_count
    
    def read_file_with_encoding(self, file_path):
        """Чтение файла с автоматическим определением кодировки"""
        encodings = ['utf-8', 'cp1251', 'windows-1251', 'iso-8859-1', 'cp866']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                return content, encoding
            except UnicodeDecodeError:
                continue
            except Exception:
                continue
        
        # Пробуем бинарное чтение как последний вариант
        try:
            with open(file_path, 'rb') as f:
                content = f.read().decode('utf-8', errors='ignore')
            return content, 'utf-8'
        except Exception as e:
            return None, None
    
    def write_file_with_encoding(self, file_path, content, encoding):
        """Запись файла с указанной кодировкой"""
        try:
            with open(file_path, 'w', encoding=encoding) as f:
                f.write(content)
        except Exception:
            # Резервный вариант с UTF-8
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
    
    def replace_commas_with_dots(self, content):
        """Замена запятых на точки в содержимом файла"""
        # Заменяем все запятые на точки
        new_content = content.replace(',', '.')
        
        # Подсчитываем количество замен
        replacements = new_content.count('.') - content.count('.') + content.count(',')
        
        return new_content, replacements

def get_plugin_class():
    return CommaReplacerPlugin