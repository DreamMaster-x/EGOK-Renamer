@echo off
chcp 65001
title Сборка EGOK Renamer

echo ========================================
echo    Сборка EGOK_Renamer в EXE
echo ========================================
echo.

echo Проверка установки PyInstaller...
pip install pyinstaller

if errorlevel 1 (
    echo Ошибка установки PyInstaller!
    pause
    exit /b 1
)

echo.
echo Сборка EXE файла...

:: Переходим в текущую директорию скрипта
cd /d "%~dp0"

:: Проверяем существование main.py
if not exist "main.py" (
    echo Ошибка: Файл main.py не найден!
    echo Убедитесь, что bat-файл находится в одной папке с main.py
    pause
    exit /b 1
)

:: Создаем папку plugins если её нет
if not exist "plugins" (
    echo Создаем папку plugins...
    mkdir plugins
    echo Создаем файл инициализации плагинов...
    echo # Пакет плагинов для EGOK Renamer > plugins\__init__.py
)

:: Проверяем наличие плагинов и создаем если нет
if not exist "plugins\example_plugin.py" (
    echo Создаем пример плагина...
    type nul > plugins\example_plugin.py
)

if not exist "plugins\file_generator_plugin.py" (
    echo Создаем плагин генератора файлов...
    type nul > plugins\file_generator_plugin.py
)

:: Проверяем наличие иконки
if exist "icon.ico" (
    echo Иконка найдена, добавляем в сборку...
    set ICON_OPTION=--icon=icon.ico
) else (
    echo Внимание: Файл иконки icon.ico не найден!
    echo Сборка продолжится без пользовательской иконки.
    set ICON_OPTION=
)

echo.
echo Начинаем сборку EXE...
echo.

:: Собираем EXE с правильными путями
pyinstaller --onefile --windowed --name "EGOK_Renamer" %ICON_OPTION% ^
--add-data "background.png;." ^
--add-data "settings.json;." ^
--add-data "icon.ico;." ^
--add-data "plugins\*;plugins" ^
--hidden-import=watchdog.observers ^
--hidden-import=watchdog.events ^
--hidden-import=PIL ^
--hidden-import=PIL._tkinter_finder ^
--hidden-import=PIL.Image ^
--hidden-import=PIL.ImageDraw ^
--hidden-import=PIL.ImageFont ^
--hidden-import=threading ^
--hidden-import=queue ^
--hidden-import=pathlib ^
--hidden-import=re ^
--hidden-import=importlib ^
--hidden-import=inspect ^
--hidden-import=importlib.util ^
--hidden-import=importlib.machinery ^
--hidden-import=requests ^
--hidden-import=json ^
--hidden-import=tksheet ^
--hidden-import=tksheet._tksheet ^
--hidden-import=tksheet._tksheet_formatters ^
--hidden-import=tksheet._tksheet_other ^
--hidden-import=tksheet._tksheet_main_table ^
--hidden-import=tksheet._tksheet_top_left_rectangle ^
--hidden-import=tksheet._tksheet_row_index ^
--hidden-import=tksheet._tksheet_header ^
--hidden-import=tksheet._tksheet_column_drag_and_drop ^
--collect-all=plugins ^
--collect-all=tksheet ^
--collect-all=PIL ^
main.py

if errorlevel 1 (
    echo.
    echo Ошибка сборки EXЕ!
    echo.
    echo Возможные решения:
    echo 1. Убедитесь что файл main.py существует
    echo 2. Проверьте установку Python и библиотек
    echo 3. Запустите install.bat перед сборкой
    echo 4. Убедитесь, что в пути к проекту нет русских букв или специальных символов
    pause
    exit /b 1
)

echo.
echo ========================================
echo    Сборка завершена успешно!
echo ========================================
echo.
echo EXE файл: dist\EGOK_Renamer.exe
echo.
echo Для запуска программы скопируйте из папки dist:
echo - EGOK_Renamer.exe
echo - background.png (рядом с EXE)
echo - settings.json (рядом с EXE)
echo - icon.ico (рядом с EXE, если используется)
echo - plugins/ (папка с плагинами)
echo.
echo НОВЫЕ ФУНКЦИИ:
echo 1. ПЛАГИН ГЕНЕРАТОРА ФАЙЛОВ - создает тестовые файлы
echo 2. Требует установленный Pillow (PIL) для генерации изображений
echo.
pause