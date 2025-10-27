#!/bin/bash
echo "Creando entorno virtual..."
python3 -m venv venv

echo "Activando entorno virtual..."
source venv/bin/activate

echo "Instalando dependencias..."
pip install -r requirements.txt

echo ""
echo "========================================"
echo "Entorno virtual creado exitosamente"
echo "========================================"
echo ""
echo "Para activar el entorno virtual, ejecuta:"
echo "  source venv/bin/activate"
echo ""
echo "Luego puedes ejecutar:"
echo "  python gui.py"
echo ""


