# Sistema de Extracción de Pagos WhatsApp → Excel

Sistema con interfaz gráfica que procesa archivos de chat de WhatsApp y extrae información de pagos automáticamente.

## Instalación

```bash
# Crear entorno virtual
python -m venv venv

# Activar entorno
venv\Scripts\activate

# Instalar dependencias
pip install -r requirements.txt
```

## Uso

### Interfaz Gráfica (Recomendado)

```bash
python gui.py
```

O simplemente haz doble clic en `iniciar_gui.bat` (Windows)

La interfaz minimalista te permite:
- **Seleccionar Archivos**: Haz clic en la zona de selección para elegir archivos
- **Procesar Archivos**: Procesa todos los archivos en la carpeta input/
- **Generar Excel**: Exporta todos los datos de la base de datos a Excel
- **Abrir Excel**: Abre el archivo Excel generado automáticamente
- **Ver Estadísticas**: Footer con totales simples (Registros, Pagos, Ahorros)
- **Log Simple**: Mensajes con íconos para seguir el proceso

### Línea de Comandos (Alternativa)

```bash
python main.py

# Opciones:
# 1. Monitoreo automático - Detecta y procesa archivos nuevos
# 2. Generar Excel actualizado - Exporta todos los datos
# 3. Procesar archivos pendientes - Procesa todo en input/
```

## Cómo Exportar Chat de WhatsApp

1. Abre WhatsApp en tu teléfono
2. Ve al chat del grupo
3. Toca el menú (⋮) → Más → Exportar chat
4. Selecciona "Sin archivos multimedia"
5. Guarda el archivo .txt en la carpeta input/ del sistema

## Archivos Generados

- `output/pagos.xlsx` - Excel con todos los pagos y resumen
- `database/pagos.db` - Base de datos SQLite
- `logs/procesamiento.log` - Registro de operaciones

## Configuración

Edita `config.json` para personalizar:
- Cortes horarios (matutino, vespertino, tarde)
- Rutas de carpetas
- Intervalo de monitoreo