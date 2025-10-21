@echo off
echo ========================================
echo   Sistema de Extracción de Pagos
echo   WhatsApp → Excel
echo ========================================
echo.
echo Iniciando interfaz gráfica...
echo.

cd /d "%~dp0"

if exist "venv\Scripts\python.exe" (
    venv\Scripts\python.exe gui.py
) else (
    python gui.py
)

if errorlevel 1 (
    echo.
    echo Error al iniciar la aplicación
    pause
)

