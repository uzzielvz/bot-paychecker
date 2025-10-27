@echo off
echo Iniciando Sistema de Gestion de Pagos...
echo.
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
    python gui.py
) else (
    echo Ejecutando sin entorno virtual...
    python gui.py
)


