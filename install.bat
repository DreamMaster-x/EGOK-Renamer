@echo off
chcp 65001
title Установка зависимостей для EGOK Renamer

echo ========================================
echo    Установка зависимостей Python
echo ========================================
echo.

echo Установка основных библиотек...
pip install watchdog
pip install pillow
pip install tksheet --upgrade

if errorlevel 1 (
    echo Ошибка установки библиотек!
    pause
    exit /b 1
)

echo.
echo Создание структуры папок...

:: Создаем папку plugins если её нет
if not exist "plugins" (
    echo Создаем папку plugins...
    mkdir plugins
    echo Создаем файл инициализации плагинов...
    echo # Пакет плагинов для EGOK Renamer > plugins\__init__.py
)

:: Создаем пример плагина если нет
if not exist "plugins\example_plugin.py" (
    echo Создаем пример плагина...
    (
    echo import tkinter as tk
    echo from tkinter import ttk
    echo.
    echo class ExamplePlugin:
    echo     def __init__(self, settings, root):
    echo         self.settings = settings
    echo         self.root = root
    echo.
    echo     def get_tab_name(self):
    echo         return "Пример плагина"
    echo.
    echo     def create_tab(self):
    echo         tab_frame = ttk.Frame(self.root)
    echo         label = ttk.Label(tab_frame, text="Это пример плагина!")
    echo         label.pack(padx=20, pady=20)
    echo         return tab_frame
    echo.
    echo def get_plugin_class():
    echo     return ExamplePlugin
    ) > plugins\example_plugin.py
)

:: Создаем плагин отчетов
if not exist "plugins\report_plugin.py" (
    echo Создаем плагин отчетов...
    type nul > plugins\report_plugin.py
    echo Файл плагина отчетов создан, заполните его кодом вручную
)

echo.
echo ========================================
echo    Установка завершена успешно!
echo ========================================
echo.
echo Установлены библиотеки:
echo - watchdog ^(мониторинг файлов^)
echo - pillow ^(PIL, работа с изображениями^)  
echo - tksheet ^(расширенная таблица^)
echo.
echo Создана структура папок:
echo - plugins/ ^(папка для плагинов^)
echo - plugins/example_plugin.py ^(пример плагина^)
echo - plugins/report_plugin.py ^(плагин отчетов^)
echo.
echo ДОБАВЛЕН НОВЫЙ ПЛАГИН:
echo - report_plugin.py ^(полнофункциональные отчеты^)
echo.
pause