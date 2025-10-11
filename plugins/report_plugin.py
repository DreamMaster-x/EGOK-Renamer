import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
import logging
from datetime import datetime
import sqlite3

try:
    import tksheet
    TKSHEET_AVAILABLE = True
except ImportError:
    TKSHEET_AVAILABLE = False

class ReportPlugin:
    """Плагин для работы с отчетами о переименованных файлах"""
    
    def __init__(self, settings, root):
        self.settings = settings
        self.root = root
        self.db_manager = None
        self.report_data = []
        self.current_route_filter = "Все"
        self.current_date_filter = None
        
        # Заголовки колонок
        self.column_headers = ["№", "Время создания", "Маршрут", "Исходное имя файла", "Новое имя файла"]
        self.column_ids = ["number", "create_time", "route", "original_name", "new_name"]
        
        # Настройки видимости колонок
        self.column_visibility = self.settings.settings.get("column_visibility", {
            "number": True,
            "create_time": True,
            "route": True,
            "original_name": True,
            "new_name": True
        })
        
        # Порядок колонок
        self.column_order = self.settings.settings.get("column_order", self.column_ids)
        
    def get_tab_name(self):
        return "Отчеты"
    
    def create_tab(self):
        """Создание вкладки с отчетами"""
        self.tab_frame = ttk.Frame(self.root)
        self.create_report_content(self.tab_frame)
        
        # Инициализируем менеджер БД
        self.db_manager = DatabaseManager()
        
        # Загружаем историю
        self.load_report_history()
        
        return self.tab_frame
    
    def create_report_content(self, parent):
        """Создание содержимого вкладки отчетов"""
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
        
        # Фильтр по маршруту и дате
        filter_frame = ttk.Frame(report_controls_frame)
        filter_frame.pack(side=tk.RIGHT, padx=5)
        
        # Фильтр по маршруту
        route_filter_frame = ttk.Frame(filter_frame)
        route_filter_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(route_filter_frame, text="Маршрут:").pack(side=tk.LEFT, padx=2)
        
        self.route_filter_var = tk.StringVar(value="Все")
        
        # Получаем значения для фильтра из истории маршрутов отчета
        route_values = ["Все"] + self.settings.settings.get("report_route_history", [])
        
        self.route_filter_cb = ttk.Combobox(
            route_filter_frame, 
            textvariable=self.route_filter_var,
            values=route_values,
            state="readonly",
            width=10
        )
        self.route_filter_cb.pack(side=tk.LEFT, padx=2)
        self.route_filter_cb.bind('<<ComboboxSelected>>', self.on_route_filter_changed)
        
        # Фильтр по дате
        date_filter_frame = ttk.Frame(filter_frame)
        date_filter_frame.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(date_filter_frame, text="Дата:").pack(side=tk.LEFT, padx=2)
        
        # Получаем все доступные даты из базы данных
        available_dates = self.get_available_dates()
        date_values = ["Все даты"] + available_dates
        
        self.date_filter_var = tk.StringVar(value="Все даты")
        
        self.date_filter_cb = ttk.Combobox(
            date_filter_frame, 
            textvariable=self.date_filter_var,
            values=date_values,
            state="readonly",
            width=12
        )
        self.date_filter_cb.pack(side=tk.LEFT, padx=2)
        self.date_filter_cb.bind('<<ComboboxSelected>>', self.on_date_filter_changed)
        
        # Кнопка обновления списка дат
        ttk.Button(date_filter_frame, text="Обновить", command=self.update_date_filter).pack(side=tk.LEFT, padx=2)
        
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
        try:
            # Создаем таблицу с включенными индексами строк
            self.report_sheet = tksheet.Sheet(
                parent,
                show_row_index=True,
                show_header=True,
                show_x_scrollbar=True,
                show_y_scrollbar=True,
                headers=self.column_headers,
                header_height=30,
                row_index_width=50,
                frame_bg="white"
            )
            
            # Включаем все необходимые привязки для Excel-подобного поведения
            self.report_sheet.enable_bindings(
                "single_select", "toggle_select", "drag_select", "row_select", 
                "column_select", "cell_select", "arrowkeys", "tab", "ctrl_a", 
                "ctrl_c", "ctrl_v", "ctrl_x", "copy", "cut", "paste", "delete",
                "edit_cell", "right_click_popup_menu", "rc_select", 
                "rc_insert_column", "rc_delete_column", "rc_insert_row", 
                "rc_delete_row", "undo", "redo", "edit_header", "drag_and_drop_column"
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
                self.report_sheet.set_cell_alignments(align="center", cells=[(r, 0) for r in range(len(self.report_data))])
                self.report_sheet.set_cell_alignments(align="center", cells=[(r, 1) for r in range(len(self.report_data))])
                self.report_sheet.set_cell_alignments(align="center", cells=[(r, 2) for r in range(len(self.report_data))])
            
            # Упаковка таблицы
            self.report_sheet.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Создаем улучшенное контекстное меню
            self.create_enhanced_context_menu()
            
            # Привязываем улучшенное контекстное меню
            self.report_sheet.bind("<Button-3>", self.show_enhanced_context_menu)
            
            logging.info("Таблица отчета инициализирована с улучшенным выделением")
            
        except Exception as e:
            logging.error(f"Ошибка создания таблицы tksheet: {e}")
            self.create_fallback_table(parent)
    
    def create_fallback_table(self, parent):
        """Создание резервной таблицы с помощью Treeview"""
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
    
    def create_enhanced_context_menu(self):
        """Создание улучшенного контекстного меню для таблицы"""
        self.enhanced_context_menu = tk.Menu(self.report_sheet, tearoff=0)
        
        # Основные операции с выделением
        self.enhanced_context_menu.add_command(label="Копировать выделенное", command=self.copy_selected_cells)
        self.enhanced_context_menu.add_command(label="Копировать как текст", command=self.copy_as_text)
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
    
    def show_enhanced_context_menu(self, event):
        """Показать улучшенное контекстное меню"""
        try:
            self.enhanced_context_menu.post(event.x_root, event.y_root)
        except Exception as e:
            logging.error(f"Ошибка показа контекстного меню: {e}")
    
    def show_report_context_menu(self, event):
        """Показать контекстное меню для отчета"""
        self.report_context_menu.post(event.x_root, event.y_root)
    
    def copy_selected_cells(self):
        """Копировать выделенные ячейки в буфер обмена"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            try:
                if hasattr(self.report_sheet, 'ctrl_c'):
                    self.report_sheet.ctrl_c()
                elif hasattr(self.report_sheet, 'copy'):
                    self.report_sheet.copy()
                else:
                    self.copy_selected_cells_manual()
                    return
                
                selected = self.report_sheet.get_selected_cells()
                if selected:
                    rows = set()
                    cols = set()
                    for row, col in selected:
                        rows.add(row)
                        cols.add(col)
                    
                    messagebox.showinfo("Успех", f"Скопировано:\n- Ячеек: {len(selected)}\n- Строк: {len(rows)}\n- Столбцов: {len(cols)}")
                else:
                    messagebox.showinfo("Информация", "Не выделены ячейки для копирования")
                    
            except Exception as e:
                logging.error(f"Ошибка копирования ячеек: {e}")
                try:
                    self.copy_selected_cells_manual()
                except Exception as e2:
                    logging.error(f"Ошибка альтернативного копирования: {e2}")
                    messagebox.showerror("Ошибка", f"Не удалось скопировать ячейки: {e}")
        else:
            self.copy_selected_files()
    
    def copy_selected_cells_manual(self):
        """Альтернативный метод копирования выделенных ячеек"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if not selected:
                messagebox.showinfo("Информация", "Не выделены ячейки для копирования")
                return
            
            text_lines = []
            current_row = None
            current_line = []
            
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
            
            all_lines = []
            for item in selected_items:
                values = self.report_tree.item(item, "values")
                if values:
                    line = "\t".join(str(value) for value in values)
                    all_lines.append(line)
            
            if all_lines:
                self.root.clipboard_clear()
                self.root.clipboard_append("\n".join(all_lines))
                messagebox.showinfo("Успех", f"Скопировано {len(all_lines)} строк в буфер обмена")
    
    def copy_all_files(self):
        """Копировать все данные отчета в буфер обмена"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            try:
                self.report_sheet.select_all()
                
                if hasattr(self.report_sheet, 'ctrl_c'):
                    self.report_sheet.ctrl_c()
                elif hasattr(self.report_sheet, 'copy'):
                    self.report_sheet.copy()
                else:
                    self.copy_all_files_manual()
                    return
                
                self.report_sheet.deselect("all")
                messagebox.showinfo("Успех", "Вся таблица скопирована в буфер обмена")
            except Exception as e:
                logging.error(f"Ошибка копирования таблицы: {e}")
                try:
                    self.copy_all_files_manual()
                except Exception as e2:
                    logging.error(f"Ошибка альтернативного копирования всей таблицы: {e2}")
                    messagebox.showerror("Ошибка", f"Не удалось скопировать таблицу: {e}")
        elif hasattr(self, 'report_tree'):
            all_items = self.report_tree.get_children()
            if not all_items:
                messagebox.showwarning("Внимание", "В отчете нет данных")
                return
            
            all_lines = []
            for item in all_items:
                values = self.report_tree.item(item, "values")
                if values:
                    line = "\t".join(str(value) for value in values)
                    all_lines.append(line)
            
            if all_lines:
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
            
            # Очищаем базу данных
            if self.db_manager:
                self.db_manager.clear_all_records()
            
            # Сбрасываем фильтры
            self.route_filter_var.set("Все")
            self.date_filter_var.set("Все даты")
            self.current_route_filter = "Все"
            self.current_date_filter = None
            
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
                    f.write(f"Фильтр по дате: {self.current_date_filter or 'Все даты'}\n")
                    f.write("=" * 80 + "\n")
                    f.write("№\tВремя\tМаршрут\tИсходное имя\tНовое имя\n")
                    f.write("-" * 80 + "\n")
                    
                    if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
                        for row in self.report_sheet.get_sheet_data():
                            if row and any(cell is not None for cell in row):
                                f.write(f"{row[0] or ''}\t{row[1] or ''}\t{row[2] or ''}\t{row[3] or ''}\t{row[4] or ''}\n")
                    else:
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
        if create_time is None:
            try:
                create_time = datetime.now().strftime('%H:%M:%S')
            except:
                create_time = datetime.now().strftime('%H:%M:%S')
        
        route = self.settings.settings["route"]
        
        # Добавляем маршрут в историю для фильтра
        if route not in self.settings.settings.get("report_route_history", []):
            self.settings.add_to_route_history(route)
            self.update_route_filter_combobox()
        
        # Номер строки
        number = len(self.report_data) + 1
        
        # Создаем строку данных
        row_data = [number, create_time, route, original_name, new_name]
        
        # Добавляем в данные отчета
        self.report_data.append(row_data)
        
        # Сохраняем в базу данных
        if self.db_manager:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.db_manager.add_record(timestamp, route, original_name, new_name, filepath)
        
        # Добавляем в таблицу
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            self.report_sheet.set_sheet_data(self.report_data)
        elif hasattr(self, 'report_tree'):
            values = (number, create_time, route, original_name, new_name)
            item_id = self.report_tree.insert("", tk.END, values=values)
            self.report_tree.see(item_id)
        
        logging.info(f"Добавлено в отчет: {original_name} -> {new_name}")
    
    def load_report_history(self):
        """Загрузка истории переименований из базы данных"""
        if not self.db_manager:
            return
            
        try:
            records = self.db_manager.get_records_by_date()
            
            self.report_data = []
            for i, record in enumerate(records):
                row_data = [
                    i + 1,
                    record['timestamp'].split(' ')[1] if ' ' in record['timestamp'] else record['timestamp'],
                    record['route'],
                    record['original_name'],
                    record['new_name']
                ]
                self.report_data.append(row_data)
            
            if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
                self.report_sheet.set_sheet_data(self.report_data)
            elif hasattr(self, 'report_tree'):
                for item in self.report_tree.get_children():
                    self.report_tree.delete(item)
                
                for row in self.report_data:
                    self.report_tree.insert("", tk.END, values=tuple(row))
            
            self.update_date_filter()
            
            logging.info(f"Загружено {len(records)} записей из истории")
        except Exception as e:
            logging.error(f"Ошибка загрузки истории отчета: {e}")
    
    def get_available_dates(self):
        """Получение всех доступных дат из базы данных"""
        if self.db_manager:
            return self.db_manager.get_all_dates()
        return []
    
    def update_route_filter_combobox(self):
        """Обновление комбобокса фильтра по маршруту"""
        route_values = ["Все"] + self.settings.settings.get("report_route_history", [])
        self.route_filter_cb['values'] = route_values
    
    def update_date_filter(self):
        """Обновление комбобокса фильтра по дате"""
        available_dates = self.get_available_dates()
        date_values = ["Все даты"] + available_dates
        self.date_filter_cb['values'] = date_values
    
    def apply_column_visibility(self):
        """Применить настройки видимости колонок"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            columns_to_show = []
            
            column_id_to_index = {column_id: i for i, column_id in enumerate(self.column_ids)}
            
            for column_id in self.column_order:
                if self.column_visibility.get(column_id, True):
                    if column_id in column_id_to_index:
                        columns_to_show.append(column_id_to_index[column_id])
            
            try:
                self.report_sheet.visible_columns = columns_to_show
                self.report_sheet.set_sheet_data(self.report_data)
            except AttributeError:
                self.report_sheet.display_columns(columns_to_show)
                self.report_sheet.set_sheet_data(self.report_data)
                
        elif hasattr(self, 'report_tree'):
            visible_columns = [col for col in self.column_order if self.column_visibility.get(col, True)]
            self.report_tree["displaycolumns"] = visible_columns
            
            column_names = {
                "number": "№",
                "create_time": "Время",
                "route": "Маршрут",
                "original_name": "Исходное имя",
                "new_name": "Новое имя"
            }
            
            for column in visible_columns:
                self.report_tree.heading(column, text=column_names.get(column, column))
    
    def on_route_filter_changed(self, event=None):
        """Обработка изменения фильтра по маршруту"""
        selected_route = self.route_filter_var.get()
        self.current_route_filter = selected_route
        self.apply_filters()
    
    def on_date_filter_changed(self, event=None):
        """Обработка изменения фильтра по дате"""
        selected_date = self.date_filter_var.get()
        self.current_date_filter = selected_date if selected_date != "Все даты" else None
        self.apply_filters()
    
    def apply_filters(self):
        """Применение всех активных фильтров"""
        if TKSHEET_AVAILABLE and hasattr(self, 'report_sheet'):
            filtered_data = []
            
            for row in self.report_data:
                route_match = (self.current_route_filter == "Все" or 
                              (len(row) > 2 and row[2] == self.current_route_filter))
                
                if route_match:
                    filtered_data.append(row)
            
            self.report_sheet.set_sheet_data(filtered_data)
        elif hasattr(self, 'report_tree'):
            all_items = self.report_tree.get_children()
            
            for item in all_items:
                values = self.report_tree.item(item, "values")
                if len(values) > 2:
                    route = values[2]
                    route_match = (self.current_route_filter == "Все" or route == self.current_route_filter)
                    
                    if route_match:
                        self.report_tree.attach(item, '', 'end')
                    else:
                        self.report_tree.detach(item)
        
        logging.info(f"Применены фильтры: маршрут={self.current_route_filter}, дата={self.current_date_filter or 'Все даты'}")
    
    def show_column_management_dialog(self):
        """Диалог управления колонками"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Управление колонками отчета")
        dialog.geometry("400x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Управление колонками отчета:", font=('Arial', 10, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        columns_frame = ttk.LabelFrame(main_frame, text="Видимые колонки (перетащите для изменения порядка)")
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.column_vars = {}
        self.column_listbox = tk.Listbox(columns_frame, selectmode=tk.SINGLE)
        scrollbar = ttk.Scrollbar(columns_frame, orient=tk.VERTICAL, command=self.column_listbox.yview)
        self.column_listbox.configure(yscrollcommand=scrollbar.set)
        
        column_names = {
            "number": "№",
            "create_time": "Время создания",
            "route": "Маршрут",
            "original_name": "Исходное имя файла",
            "new_name": "Новое имя файла"
        }
        
        for column_id in self.column_order:
            if column_id in self.column_visibility:
                display_name = column_names.get(column_id, column_id)
                status = "✓" if self.column_visibility[column_id] else "✗"
                self.column_listbox.insert(tk.END, f"{status} {display_name}")
                self.column_vars[column_id] = self.column_visibility[column_id]
        
        self.column_listbox.bind("<Double-Button-1>", self.toggle_column_visibility)
        self.column_listbox.bind('<ButtonPress-1>', self.on_drag_start)
        self.column_listbox.bind('<B1-Motion>', self.on_drag_motion)
        self.column_listbox.bind('<ButtonRelease-1>', self.on_drag_release)
        
        self.column_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(buttons_frame, text="Сохранить", command=lambda: self.save_column_settings(dialog)).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Сбросить", command=self.reset_column_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Отмена", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def toggle_column_visibility(self, event):
        """Переключение видимости колонки при двойном клике"""
        selection = self.column_listbox.curselection()
        if selection:
            index = selection[0]
            column_id = self.column_order[index]
            self.column_vars[column_id] = not self.column_vars[column_id]
            
            column_names = {
                "number": "№",
                "create_time": "Время создания",
                "route": "Маршрут",
                "original_name": "Исходное имя файла",
                "new_name": "Новое имя файла"
            }
            
            display_name = column_names.get(column_id, column_id)
            status = "✓" if self.column_vars[column_id] else "✗"
            self.column_listbox.delete(index)
            self.column_listbox.insert(index, f"{status} {display_name}")
            self.column_listbox.selection_set(index)
    
    def on_drag_start(self, event):
        """Начало перетаскивания элемента списка"""
        self.drag_start_index = self.column_listbox.nearest(event.y)
    
    def on_drag_motion(self, event):
        """Перетаскивание элемента списка"""
        pass
    
    def on_drag_release(self, event):
        """Завершение перетаскивания элемента списка"""
        end_index = self.column_listbox.nearest(event.y)
        if hasattr(self, 'drag_start_index') and self.drag_start_index != end_index:
            items = list(self.column_vars.items())
            item_to_move = items.pop(self.drag_start_index)
            items.insert(end_index, item_to_move)
            
            self.column_order = [item[0] for item in items]
            self.column_vars = dict(items)
            
            self.column_listbox.delete(0, tk.END)
            
            column_names = {
                "number": "№",
                "create_time": "Время создания",
                "route": "Маршрут",
                "original_name": "Исходное имя файла",
                "new_name": "Новое имя файла"
            }
            
            for column_id, is_visible in items:
                display_name = column_names.get(column_id, column_id)
                status = "✓" if is_visible else "✗"
                self.column_listbox.insert(tk.END, f"{status} {display_name}")
            
            self.column_listbox.selection_set(end_index)
    
    def save_column_settings(self, dialog):
        """Сохранение настроек колонок"""
        self.column_visibility = self.column_vars.copy()
        self.settings.update_setting("column_visibility", self.column_visibility)
        self.settings.update_setting("column_order", self.column_order)
        
        self.apply_column_visibility()
        dialog.destroy()
        logging.info("Настройки колонок сохранены")
    
    def reset_column_settings(self):
        """Сброс настроек колонок к значениям по умолчанию"""
        self.column_visibility = {
            "number": True,
            "create_time": True,
            "route": True,
            "original_name": True,
            "new_name": True
        }
        self.column_order = ["number", "create_time", "route", "original_name", "new_name"]
        
        self.column_listbox.delete(0, tk.END)
        
        column_names = {
            "number": "№",
            "create_time": "Время создания",
            "route": "Маршрут",
            "original_name": "Исходное имя файла",
            "new_name": "Новое имя файла"
        }
        
        for column_id in self.column_order:
            display_name = column_names.get(column_id, column_id)
            status = "✓" if self.column_visibility[column_id] else "✗"
            self.column_listbox.insert(tk.END, f"{status} {display_name}")
        
        logging.info("Настройки колонок сброшены к значениям по умолчанию")
    
    # Методы для улучшенного контекстного меню
    def copy_as_text(self):
        """Копирование выделенного как форматированный текст"""
        try:
            selected = self.report_sheet.get_selected_cells()
            if not selected:
                messagebox.showinfo("Информация", "Не выделены ячейки для копирования")
                return
            
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
                filetypes=[("CSV files", "*.csv"), ("Text files", "*.txt"), ("All files", "*.*")],
                title="Экспорт выделенных данных"
            )
            
            if file_path:
                data_to_export = []
                
                rows_data = {}
                for row, col in selected:
                    if row not in rows_data:
                        rows_data[row] = {}
                    cell_value = self.report_sheet.get_cell_data(row, col)
                    rows_data[row][col] = cell_value if cell_value is not None else ""
                
                for row in sorted(rows_data.keys()):
                    row_data = []
                    for col in sorted(rows_data[row].keys()):
                        row_data.append(str(rows_data[row][col]))
                    data_to_export.append(",".join(row_data))
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write("\n".join(data_to_export))
                
                messagebox.showinfo("Успех", f"Данные экспортированы в:\n{file_path}")
                
        except Exception as e:
            logging.error(f"Ошибка экспорта выделенных данных: {e}")
            messagebox.showerror("Ошибка", f"Не удалось экспортировать данные: {e}")

def get_plugin_class():
    return ReportPlugin


class DatabaseManager:
    """Менеджер базы данных для хранения истории переименований"""
    
    def __init__(self, db_file="rename_history.db"):
        self.db_file = db_file
        self.init_database()
    
    def init_database(self):
        """Инициализация базы данных"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS rename_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TEXT NOT NULL,
                    create_date TEXT NOT NULL,
                    route TEXT NOT NULL,
                    original_name TEXT NOT NULL,
                    new_name TEXT NOT NULL,
                    file_path TEXT NOT NULL
                )
            ''')
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_date ON rename_history(create_date)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_route ON rename_history(route)')
            
            conn.commit()
            conn.close()
            logging.info("База данных инициализирована успешно")
            
        except Exception as e:
            logging.error(f"Ошибка инициализации базы данных: {e}")
    
    def add_record(self, timestamp, route, original_name, new_name, file_path):
        """Добавление записи в базу данных"""
        try:
            create_date = datetime.now().strftime("%Y-%m-%d")
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO rename_history 
                (timestamp, create_date, route, original_name, new_name, file_path)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (timestamp, create_date, route, original_name, new_name, file_path))
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logging.error(f"Ошибка добавления записи в базу данных: {e}")
            return False
    
    def get_records_by_date(self, target_date=None):
        """Получение записей за определенную дату"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            if target_date:
                cursor.execute('SELECT * FROM rename_history WHERE create_date = ? ORDER BY timestamp DESC', (target_date,))
            else:
                cursor.execute('SELECT * FROM rename_history ORDER BY timestamp DESC')
            
            records = cursor.fetchall()
            conn.close()
            
            columns = ['id', 'timestamp', 'create_date', 'route', 'original_name', 'new_name', 'file_path']
            result = []
            for record in records:
                result.append(dict(zip(columns, record)))
            
            return result
        except Exception as e:
            logging.error(f"Ошибка получения записей из базы данных: {e}")
            return []
    
    def get_all_dates(self):
        """Получение всех уникальных дат из базы данных"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('SELECT DISTINCT create_date FROM rename_history ORDER BY create_date DESC')
            
            dates = [row[0] for row in cursor.fetchall()]
            conn.close()
            return dates
        except Exception as e:
            logging.error(f"Ошибка получения дат из базы данных: {e}")
            return []
    
    def clear_all_records(self):
        """Удаление всех записей"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM rename_history')
            
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            logging.error(f"Ошибка очистки базы данных: {e}")
            return False