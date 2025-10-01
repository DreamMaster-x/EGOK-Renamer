# Пример плагина для EGOK Renamer
import tkinter as tk
from tkinter import ttk, messagebox
import logging

class ExamplePlugin:
    """Пример плагина - демонстрирует возможности расширения"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
    
    def get_tab_name(self):
        """Возвращает название вкладки"""
        return "Пример плагина"
    
    def create_tab(self):
        """Создает содержимое вкладки"""
        try:
            frame = ttk.Frame(self.root)
            
            # Заголовок
            title_label = ttk.Label(frame, text="Демонстрация системы плагинов", 
                                   font=('Arial', 12, 'bold'))
            title_label.pack(pady=10)
            
            # Описание
            desc_text = """Это пример плагина, демонстрирующий возможности расширения программы.

Разработчики могут создавать собственные плагины:
1. Создать файл в папке plugins/
2. Унаследоваться от BasePlugin
3. Реализовать методы get_tab_name() и create_tab()

Примеры возможных плагинов:
- Работа с картами
- Обработка видео
- Выгрузка в облако
- Анализ данных
- Интеграция с другими системами"""
            
            desc_label = ttk.Label(frame, text=desc_text, justify=tk.LEFT)
            desc_label.pack(pady=10, padx=10, fill=tk.X)
            
            # Разделитель
            separator = ttk.Separator(frame, orient='horizontal')
            separator.pack(fill=tk.X, padx=10, pady=10)
            
            # Демонстрационные элементы
            demo_frame = ttk.LabelFrame(frame, text="Демо-элементы")
            demo_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # Кнопка
            ttk.Button(demo_frame, text="Тестовая кнопка", 
                      command=self.show_message).pack(pady=5)
            
            # Поле ввода
            entry_var = tk.StringVar(value="Пример текста")
            entry = ttk.Entry(demo_frame, textvariable=entry_var, width=30)
            entry.pack(pady=5)
            
            # Combobox
            combo_var = tk.StringVar(value="Вариант 1")
            combo = ttk.Combobox(demo_frame, textvariable=combo_var,
                                values=["Вариант 1", "Вариант 2", "Вариант 3"])
            combo.pack(pady=5)
            
            # Прогресс-бар
            progress = ttk.Progressbar(demo_frame, mode='indeterminate')
            progress.pack(fill=tk.X, padx=5, pady=5)
            
            # Кнопка для прогресс-бара
            def toggle_progress():
                if progress.cget('mode') == 'indeterminate':
                    progress.start(10)
                else:
                    progress.stop()
            
            ttk.Button(demo_frame, text="Запустить/Остановить прогресс", 
                      command=toggle_progress).pack(pady=5)
            
            logging.info("Плагин 'Пример' успешно загружен")
            return frame
            
        except Exception as e:
            logging.error(f"Ошибка создания вкладки плагина: {e}")
            return None
    
    def show_message(self):
        """Показать тестовое сообщение"""
        messagebox.showinfo("Тест плагина", 
                           "Это сообщение из примера плагина!\n"
                           "Плагин успешно интегрирован в основное приложение.")