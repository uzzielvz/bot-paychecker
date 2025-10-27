# Sistema de Gestión de Pagos desde WhatsApp

Sistema automatizado para extraer y gestionar registros de pagos desde chats de WhatsApp (.txt) y actualizar un archivo Excel.

## Características

- Extracción automática de pagos desde chats de WhatsApp
- Normalización de datos (nombres, sucursales, montos)
- Gestión de confirmaciones
- Columna de corte (Matutino/Vespertino)
- Mapeo de IDs de grupos con sus nombres
- Interfaz gráfica
- Drag-and-drop de archivos
- Prevención de duplicados por timestamp

## Requisitos

- Python 3.x
- pandas
- openpyxl
- tkinterdnd2

## Instalación

### Windows

```bash
setup_venv.bat
iniciar.bat
```

### Linux/Mac

```bash
bash setup_venv.sh
bash iniciar.sh
```

### Instalación Global

```bash
pip install -r requirements.txt
python gui.py
```

## Uso

### Modo GUI

```bash
venv\Scripts\activate
python gui.py
```

### Modo Línea de Comandos

```bash
python payment_manager.py
```

## Estructura del Proyecto

- payment_manager.py - Lógica principal
- gui.py - Interfaz gráfica
- config.json - Configuración
- Pagos.xlsx - Registros
- log.txt - Log del sistema

## Configuración

### config.json

- horarios: Define horario matutino y vespertino
- archivo_procesado: Timestamp del último archivo procesado
- mapeo_id_grupos: Mapea IDs a nombres y sucursales

### Formato de Entrada

```
[24/10/25, 10:51:52] Uzziel: Grupo BIENVENIDOS 
ID 000094
Pago 12921
Ahorro 1293 
Sucursal Ixtapaluca
```

## Funcionalidades GUI

1. Subir Pagos: Seleccionar o arrastrar archivos .txt de pagos
2. Procesar Pagos: Extraer y guardar registros en Excel
3. Subir Confirmaciones: Seleccionar archivos con confirmaciones
4. Procesar Confirmaciones: Marcar como confirmados
5. Ver Excel: Abrir archivo Pagos.xlsx
6. Limpiar Registros: Eliminar todos los datos

## Columnas del Excel

- ID
- Grupo
- Fecha
- Hora
- Pago
- Ahorro
- Total
- Número de Pago
- Sucursal
- Corte
- Confirmado

## Hojas del Excel

- Pagos: Todos los registros
- Pagos Confirmados: Pagos confirmados
- Meta: Información interna (oculta)

## Notas

- Los nombres se normalizan a mayúsculas
- Las sucursales se normalizan sin acentos
- Prevención de duplicados por timestamp
- Los IDs se pueden agregar manualmente al config.json
