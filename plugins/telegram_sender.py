# Плагин для отправки файлов в Telegram канал
import tkinter as tk
from tkinter import ttk, messagebox
import logging
import threading
import time
import os
from pathlib import Path
import requests
import json
from datetime import datetime

class TelegramSenderPlugin:
    """Плагин для отправки файлов в Telegram канал с задержкой"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.is_monitoring = False
        self.monitor_thread = None
        self.stop_monitor = False
        self.sent_files = set()
        
        # Загрузка настроек плагина
        self.plugin_settings = self.settings.settings.get("telegram_sender", {})
        
    def get_tab_name(self):
        """Возвращает название вкладки"""
        return "Telegram Отправка"
    
    def create_tab(self):
        """Создает содержимое вкладки"""
        try:
            frame = ttk.Frame(self.root)
            
            # Заголовок
            title_label = ttk.Label(frame, text="Отправка файлов в Telegram канал", 
                                   font=('Arial', 12, 'bold'))
            title_label.pack(pady=10)
            
            # Основные настройки
            settings_frame = ttk.LabelFrame(frame, text="Настройки Telegram")
            settings_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Токен бота
            token_frame = ttk.Frame(settings_frame)
            token_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Label(token_frame, text="Токен бота:").pack(side=tk.LEFT)
            self.bot_token_var = tk.StringVar(value=self.plugin_settings.get("bot_token", ""))
            token_entry = ttk.Entry(token_frame, textvariable=self.bot_token_var, width=40, show="*")
            token_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # ID канала
            channel_frame = ttk.Frame(settings_frame)
            channel_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Label(channel_frame, text="ID канала:").pack(side=tk.LEFT)
            self.channel_id_var = tk.StringVar(value=self.plugin_settings.get("channel_id", ""))
            channel_entry = ttk.Entry(channel_frame, textvariable=self.channel_id_var, width=40)
            channel_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # Задержка отправки
            delay_frame = ttk.Frame(settings_frame)
            delay_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Label(delay_frame, text="Задержка (секунды):").pack(side=tk.LEFT)
            self.delay_var = tk.StringVar(value=str(self.plugin_settings.get("delay_seconds", 10)))
            delay_entry = ttk.Entry(delay_frame, textvariable=self.delay_var, width=10)
            delay_entry.pack(side=tk.LEFT, padx=5)
            
            # Папка для мониторинга
            folder_frame = ttk.Frame(settings_frame)
            folder_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Label(folder_frame, text="Папка мониторинга:").pack(side=tk.LEFT)
            self.monitor_folder_var = tk.StringVar(
                value=self.plugin_settings.get("monitor_folder", self.settings.settings["folder"])
            )
            folder_entry = ttk.Entry(folder_frame, textvariable=self.monitor_folder_var, width=30)
            folder_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            ttk.Button(folder_frame, text="Обзор", command=self.browse_monitor_folder).pack(side=tk.LEFT)
            
            # Расширения файлов
            ext_frame = ttk.Frame(settings_frame)
            ext_frame.pack(fill=tk.X, padx=5, pady=2)
            ttk.Label(ext_frame, text="Расширения файлов:").pack(side=tk.LEFT)
            self.extensions_var = tk.StringVar(
                value=self.plugin_settings.get("extensions", self.settings.settings["extensions"])
            )
            ext_entry = ttk.Entry(ext_frame, textvariable=self.extensions_var, width=30)
            ext_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
            
            # Разделитель
            separator = ttk.Separator(frame, orient='horizontal')
            separator.pack(fill=tk.X, padx=10, pady=10)
            
            # Управление
            control_frame = ttk.LabelFrame(frame, text="Управление")
            control_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Кнопки управления
            button_frame = ttk.Frame(control_frame)
            button_frame.pack(pady=5)
            
            self.monitor_button = ttk.Button(
                button_frame, 
                text="Запуск мониторинга", 
                command=self.toggle_monitoring,
                width=20
            )
            self.monitor_button.pack(side=tk.LEFT, padx=5)
            
            ttk.Button(
                button_frame, 
                text="Тест отправки", 
                command=self.test_send,
                width=15
            ).pack(side=tk.LEFT, padx=5)
            
            ttk.Button(
                button_frame, 
                text="Сохранить настройки", 
                command=self.save_settings,
                width=15
            ).pack(side=tk.LEFT, padx=5)
            
            # Статус
            self.status_var = tk.StringVar(value="Мониторинг остановлен")
            status_label = ttk.Label(control_frame, textvariable=self.status_var, foreground="red")
            status_label.pack(pady=5)
            
            # Лог отправки
            log_frame = ttk.LabelFrame(frame, text="Лог отправки")
            log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
            
            self.log_text = tk.Text(log_frame, wrap=tk.WORD, height=8, state=tk.DISABLED)
            scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
            self.log_text.configure(yscrollcommand=scrollbar.set)
            
            self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            # Обновляем состояние кнопки
            self.update_monitor_button()
            
            logging.info("Плагин 'Telegram Отправка' успешно загружен")
            return frame
            
        except Exception as e:
            logging.error(f"Ошибка создания вкладки плагина: {e}")
            return None
    
    def browse_monitor_folder(self):
        """Выбор папки для мониторинга"""
        folder = tk.filedialog.askdirectory()
        if folder:
            self.monitor_folder_var.set(folder)
    
    def save_settings(self):
        """Сохранение настроек плагина"""
        try:
            self.plugin_settings = {
                "bot_token": self.bot_token_var.get(),
                "channel_id": self.channel_id_var.get(),
                "delay_seconds": int(self.delay_var.get()),
                "monitor_folder": self.monitor_folder_var.get(),
                "extensions": self.extensions_var.get()
            }
            
            self.settings.settings["telegram_sender"] = self.plugin_settings
            self.settings.save_settings()
            
            messagebox.showinfo("Успех", "Настройки плагина сохранены!")
            self.add_log("Настройки сохранены")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения настроек: {e}")
    
    def toggle_monitoring(self):
        """Переключение мониторинга"""
        if self.is_monitoring:
            self.stop_monitoring()
        else:
            self.start_monitoring()
        self.update_monitor_button()
    
    def update_monitor_button(self):
        """Обновление внешнего вида кнопки мониторинга"""
        if self.is_monitoring:
            self.monitor_button.config(text="Остановить мониторинг")
            self.status_var.set("Мониторинг активен")
        else:
            self.monitor_button.config(text="Запуск мониторинга")
            self.status_var.set("Мониторинг остановлен")
    
    def start_monitoring(self):
        """Запуск мониторинга"""
        if not self.bot_token_var.get() or not self.channel_id_var.get():
            messagebox.showerror("Ошибка", "Заполните токен бота и ID канала!")
            return
        
        if not os.path.exists(self.monitor_folder_var.get()):
            messagebox.showerror("Ошибка", "Папка для мониторинга не существует!")
            return
        
        self.is_monitoring = True
        self.stop_monitor = False
        self.monitor_thread = threading.Thread(target=self.monitor_folder)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        
        self.add_log("Мониторинг запущен")
        logging.info("Telegram мониторинг запущен")
    
    def stop_monitoring(self):
        """Остановка мониторинга"""
        self.is_monitoring = False
        self.stop_monitor = True
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)
        
        self.add_log("Мониторинг остановлен")
        logging.info("Telegram мониторинг остановлен")
    
    def monitor_folder(self):
        """Мониторинг папки на наличие новых файлов"""
        monitored_folder = self.monitor_folder_var.get()
        extensions = [ext.strip().lower() for ext in self.extensions_var.get().split(",")]
        delay_seconds = int(self.delay_var.get())
        
        self.add_log(f"Начало мониторинга папки: {monitored_folder}")
        
        # Получаем текущий список файлов
        current_files = set()
        if os.path.exists(monitored_folder):
            for filename in os.listdir(monitored_folder):
                filepath = os.path.join(monitored_folder, filename)
                if os.path.isfile(filepath):
                    file_ext = Path(filename).suffix.lower()[1:]
                    if file_ext in extensions:
                        current_files.add(filepath)
        
        while self.is_monitoring and not self.stop_monitor:
            try:
                # Проверяем новые файлы
                new_files = set()
                if os.path.exists(monitored_folder):
                    for filename in os.listdir(monitored_folder):
                        filepath = os.path.join(monitored_folder, filename)
                        if (os.path.isfile(filepath) and 
                            filepath not in self.sent_files and
                            filepath not in current_files):
                            
                            file_ext = Path(filename).suffix.lower()[1:]
                            if file_ext in extensions:
                                new_files.add(filepath)
                
                # Обрабатываем новые файлы
                for filepath in new_files:
                    if filepath not in self.sent_files:
                        # Запускаем отправку с задержкой
                        threading.Thread(
                            target=self.send_file_with_delay,
                            args=(filepath, delay_seconds),
                            daemon=True
                        ).start()
                        
                        self.add_log(f"Файл добавлен в очередь отправки: {Path(filepath).name}")
                
                # Обновляем текущий список файлов
                current_files.update(new_files)
                
                # Небольшая пауза перед следующей проверкой
                time.sleep(1)
                
            except Exception as e:
                self.add_log(f"Ошибка мониторинга: {e}")
                logging.error(f"Ошибка мониторинга Telegram: {e}")
                time.sleep(5)
    
    def send_file_with_delay(self, filepath, delay_seconds):
        """Отправка файла с задержкой"""
        try:
            filename = Path(filepath).name
            self.add_log(f"Ожидание {delay_seconds} сек. перед отправкой: {filename}")
            
            # Ожидаем указанное время
            time.sleep(delay_seconds)
            
            # Проверяем, существует ли еще файл
            if not os.path.exists(filepath):
                self.add_log(f"Файл удален, отправка отменена: {filename}")
                return
            
            # Отправляем файл
            success = self.send_to_telegram(filepath)
            
            if success:
                self.sent_files.add(filepath)
                self.add_log(f"Файл успешно отправлен: {filename}")
            else:
                self.add_log(f"Ошибка отправки файла: {filename}")
                
        except Exception as e:
            self.add_log(f"Ошибка при отправке файла: {e}")
            logging.error(f"Ошибка отправки файла в Telegram: {e}")
    
    def send_to_telegram(self, filepath):
        """Отправка файла в Telegram канал"""
        try:
            bot_token = self.bot_token_var.get()
            channel_id = self.channel_id_var.get()
            filename = Path(filepath).name
            
            # URL для отправки документа
            url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
            
            # Читаем файл
            with open(filepath, 'rb') as file:
                files = {'document': (filename, file)}
                data = {'chat_id': channel_id}
                
                # Отправляем запрос
                response = requests.post(url, files=files, data=data, timeout=30)
                
            if response.status_code == 200:
                logging.info(f"Файл отправлен в Telegram: {filename}")
                return True
            else:
                error_msg = response.json().get('description', 'Unknown error')
                self.add_log(f"Ошибка Telegram API: {error_msg}")
                logging.error(f"Ошибка отправки в Telegram: {error_msg}")
                return False
                
        except Exception as e:
            self.add_log(f"Ошибка соединения: {e}")
            logging.error(f"Ошибка соединения с Telegram: {e}")
            return False
    
    def test_send(self):
        """Тестовая отправка"""
        if not self.bot_token_var.get() or not self.channel_id_var.get():
            messagebox.showerror("Ошибка", "Заполните токен бота и ID канала!")
            return
        
        # Создаем тестовый файл
        test_filename = f"test_file_{int(time.time())}.txt"
        test_filepath = os.path.join(self.monitor_folder_var.get(), test_filename)
        
        try:
            with open(test_filepath, 'w', encoding='utf-8') as f:
                f.write(f"Тестовый файл для проверки отправки в Telegram\n")
                f.write(f"Создан: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Программа: EGOK Renamer")
            
            self.add_log(f"Создан тестовый файл: {test_filename}")
            
            # Отправляем тестовый файл
            success = self.send_to_telegram(test_filepath)
            
            if success:
                messagebox.showinfo("Успех", "Тестовая отправка прошла успешно!")
            else:
                messagebox.showerror("Ошибка", "Тестовая отправка не удалась!")
            
            # Удаляем тестовый файл
            try:
                os.remove(test_filepath)
            except:
                pass
                
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка создания тестового файла: {e}")
    
    def add_log(self, message):
        """Добавление сообщения в лог"""
        try:
            timestamp = datetime.now().strftime('%H:%M:%S')
            log_message = f"[{timestamp}] {message}\n"
            
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, log_message)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
            
        except Exception as e:
            logging.error(f"Ошибка добавления лога: {e}")