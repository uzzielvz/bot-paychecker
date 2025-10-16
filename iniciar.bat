@echo off
echo ============================================================
echo   SISTEMA DE EXTRACCION DE PAGOS WHATSAPP - EXCEL
echo ============================================================
echo.
echo Activando entorno virtual...
call venv\Scripts\activate.bat
echo.
echo Iniciando sistema...
python main.py
pause

