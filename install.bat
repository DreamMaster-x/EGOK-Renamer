@echo off
chcp 65001
title Установка EGOK Renamer

echo ========================================
echo    Установка EGOK Renamer
echo ========================================
echo.

echo Создание структуры папок...
if not exist "plugins" mkdir plugins
if not exist "plugins\__init__.py" (
    echo # Пакет плагинов для EGOK Renamer > plugins\__init__.py
)

echo Установка необходимых библиотек...
pip install watchdog pillow requests

if errorlevel 1 (
    echo Ошибка установки библиотек!
    pause
    exit /b 1
)

echo.
echo ========================================
echo    УСТАНОВКА ПЛАГИНОВ - ИНСТРУКЦИЯ
echo ========================================
echo.
echo СПОСОБ 1 (Автоматический):
echo 1. В программе: Настройки -> Установить плагин
echo 2. Выбете файл .py с плагином
echo 3. Перезапустите программу
echo.
echo СПОСОБ 2 (Вручную):
echo 1. Скопируйте файл плагина в папку plugins/
echo 2. Запустите программу
echo 3. Настройки -> Управление плагинами
echo 4. Отметьте плагин галочкой
echo 5. Сохраните и перезапустите программу
echo.
echo Установка завершена успешно!
echo.
echo Для сборки EXE файла запустите build_exe.bat
echo.
pause