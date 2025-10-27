#!/bin/bash
echo "Iniciando Sistema de Gestion de Pagos..."
echo ""
if [ -d "venv" ]; then
    source venv/bin/activate
    python gui.py
else
    echo "Ejecutando sin entorno virtual..."
    python gui.py
fi


