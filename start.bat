@echo off
chcp 65001 >nul
title Анализ рекламаций и операционного контроля
echo ============================================================
echo   Анализ рекламаций и операционного контроля
echo ============================================================
echo.

cd /d "%~dp0"

echo Проверка установки зависимостей...
python -c "import flask, openpyxl, reportlab" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Зависимости не установлены. Установка...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ОШИБКА] Не удалось установить зависимости.
        echo Убедитесь, что Python установлен: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
)

echo.
echo Запуск сервера...
echo Адрес: http://127.0.0.1:5000
echo Закройте это окно для остановки приложения.
echo.

python app.py

echo.
echo Приложение остановлено.
pause