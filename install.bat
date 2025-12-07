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
pip install pyserial

if errorlevel 1 (
    echo Ошибка установки библиотек!
    pause
    exit /b 1
)

echo.
echo Установка библиотек для плагина PDF->KML...
pip install PyPDF2
pip install simplekml

if errorlevel 1 (
    echo Ошибка установки библиотек для плагина PDF->KML!
    echo Плагин будет работать без поддержки PDF
    pause
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

:: Создаем плагин генератора файлов
if not exist "plugins\file_generator_plugin.py" (
    echo Создаем плагин генератора файлов...
    type nul > plugins\file_generator_plugin.py
    echo Файл плагина создан, заполните его кодом вручную
)

:: Создаем плагин телеметрии
if not exist "plugins\telemetry_plugin.py" (
    echo Создаем плагин обработки телеметрии...
    echo Импортируем существующий файл...
    copy "telemetry_plugin.py" "plugins\telemetry_plugin.py" >nul 2>&1
    if errorlevel 1 (
        echo Файл telemetry_plugin.py не найден, создаем пустой...
        type nul > plugins\telemetry_plugin.py
    )
    echo Файл плагина телеметрии создан
)

:: Создаем плагин PDF->KML
if not exist "plugins\pdf_kml_plugin.py" (
    echo Создаем плагин PDF->KML...
    type nul > plugins\pdf_kml_plugin.py
    echo Файл плагина PDF->KML создан
)

echo.
echo Обновление настроек для включения плагинов...
python -c "
import json
import os
try:
    if os.path.exists('settings.json'):
        with open('settings.json', 'r', encoding='utf-8') as f:
            settings = json.load(f)
        
        enabled_plugins = settings.get('enabled_plugins', [])
        if 'telemetry_plugin' not in enabled_plugins:
            enabled_plugins.append('telemetry_plugin')
            settings['enabled_plugins'] = enabled_plugins
            
            with open('settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            print('Настройки обновлены: telemetry_plugin добавлен в enabled_plugins')
        else:
            print('telemetry_plugin уже в enabled_plugins')
    else:
        print('Файл settings.json не найден, будут использованы настройки по умолчанию')
except Exception as e:
    print(f'Ошибка обновления настроек: {e}')
"

echo.
echo ========================================
echo    Установка завершена успешно!
echo ========================================
echo.
echo Установлены библиотеки:
echo - watchdog ^(мониторинг файлов^)
echo - pillow ^(PIL, работа с изображениями^)  
echo - tksheet ^(расширенная таблица^)
echo - pyserial ^(работа с COM портами^)
echo - PyPDF2 ^(чтение PDF файлов^)
echo - simplekml ^(генерация KML файлов^)
echo.
echo Создана структура папок:
echo - plugins/ ^(папка для плагинов^)
echo - plugins/example_plugin.py ^(пример плагина^)
echo - plugins/file_generator_plugin.py ^(плагин генератора^)
echo - plugins/telemetry_plugin.py ^(плагин телеметрии^)
echo - plugins/pdf_kml_plugin.py ^(плагин PDF->KML^)
echo.
echo Плагин telemetry_plugin добавлен в настройки!
echo.
echo Для работы плагина PDF->KML необходимы PyPDF2 и simplekml!
echo.
pause