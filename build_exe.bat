@echo off
chcp 65001
title Сборка EGOK Renamer

echo ========================================
echo    Сборка EGOK_Renamer в EXE
echo ========================================
echo.

echo Шаг 1: Проверка основных файлов...
if not exist "EGOK_Renamer.py" (
    echo ОШИБКА: Файл EGOK_Renamer.py не найден!
    goto error
)

echo Шаг 2: Создание структуры папок...
if not exist "plugins" mkdir plugins
if not exist "plugins\__init__.py" echo # Пакет плагинов > plugins\__init__.py

echo Шаг 3: Установка библиотек...
pip install watchdog pillow tksheet pyserial PyPDF2 simplekml --quiet

echo Шаг 4: Сборка EXE...
echo Пожалуйста, подождите...

pyinstaller --noconfirm --onefile --windowed ^
--name "EGOK_Renamer" ^
--icon=icon.ico ^
--add-data "background.png;." ^
--add-data "settings.json;." ^
--add-data "icon.ico;." ^
--add-data "plugins;plugins" ^
--hidden-import=watchdog.observers ^
--hidden-import=watchdog.events ^
--hidden-import=PIL ^
--hidden-import=PIL._tkinter_finder ^
--hidden-import=PIL.Image ^
--hidden-import=threading ^
--hidden-import=queue ^
--hidden-import=pathlib ^
--hidden-import=re ^
--hidden-import=importlib ^
--hidden-import=inspect ^
--hidden-import=json ^
--hidden-import=tksheet ^
--hidden-import=sqlite3 ^
--hidden-import=serial ^
--hidden-import=serial.tools.list_ports ^
--hidden-import=serial.serialutil ^
--hidden-import=serial.win32 ^
--hidden-import=PyPDF2 ^
--hidden-import=simplekml ^
--hidden-import=binascii ^
--hidden-import=math ^
EGOK_Renamer.py

if %errorlevel% neq 0 (
    echo.
    echo ОШИБКА при сборке!
    goto error
)

echo.
echo ========================================
echo    СБОРКА УСПЕШНО ЗАВЕРШЕНА!
echo ========================================
echo.
echo Создан файл: dist\EGOK_Renamer.exe
echo.
echo Для запуска программы:
echo 1. Откройте папку 'dist'
echo 2. Запустите 'EGOK_Renamer.exe'
echo.
echo Нажмите любую клавишу для выхода...
pause >nul
exit /b 0

:error
echo.
echo ========================================
echo    ОШИБКА СБОРКИ
echo ========================================
echo.
echo Возможные причины:
echo 1. Python не установлен
echo 2. Нет прав для записи
echo 3. Повреждены файлы проекта
echo.
echo Решения:
echo 1. Установите Python 3.8+
echo 2. Запустите от имени администратора
echo 3. Проверьте файлы в папке
echo.
echo Нажмите любую клавишу для выхода...
pause >nul
exit /b 1