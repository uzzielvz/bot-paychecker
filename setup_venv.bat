@echo off
echo Creando entorno virtual...
python -m venv venv

echo Activando entorno virtual...
call venv\Scripts\activate.bat

echo Instalando dependencias...
pip install -r requirements.txt

echo.
echo ========================================
echo Entorno virtual creado exitosamente
echo ========================================
echo.
echo Para activar el entorno virtual, ejecuta:
echo   venv\Scripts\activate
echo.
echo Luego puedes ejecutar:
echo   python gui.py
echo.
pause


