# Sistema de Gestión de Pagos desde WhatsApp

Sistema automatizado para extraer y gestionar registros de pagos desde chats de WhatsApp (.txt) y actualizar un archivo Excel.

## Características

- ✅ Extracción automática de pagos desde chats de WhatsApp
- ✅ Normalización de datos (nombres, sucursales, montos)
- ✅ Gestión de confirmaciones
- ✅ Columna de "Corte" (Matutino/Vespertino)
- ✅ Mapeo de IDs de grupos con sus nombres
- ✅ Interfaz gráfica intuitiva
- ✅ Drag-and-drop de archivos
- ✅ Prevención de duplicados por timestamp

## Requisitos

- Python 3.x
- pandas
- openpyxl
- tkinterdnd2

## Instalación

### Opción 1: Con Entorno Virtual (Recomendado)

**Windows:**
```bash
setup_venv.bat
iniciar.bat
```

**Linux/Mac:**
```bash
bash setup_venv.sh
bash iniciar.sh
```

### Opción 2: Instalación Global

```bash
pip install -r requirements.txt
python gui.py
```

## Uso

### Modo GUI (Recomendado)

```bash
# Con entorno virtual
venv\Scripts\activate    # Windows
venv/bin/activate       # Linux/Mac
python gui.py

# Sin entorno virtual
python gui.py
```

### Modo Línea de Comandos

```bash
python payment_manager.py
```

## Estructura del Proyecto

- `payment_manager.py` - Lógica principal del sistema
- `gui.py` - Interfaz gráfica
- `config.json` - Configuración (mapeo de grupos, horarios)
- `Pagos.xlsx` - Archivo Excel con los registros
- `log.txt` - Log del sistema

## Archivos de Configuración

### config.json

Contiene:
- **horarios**: Define horario matutino y vespertino
- **archivo_procesado**: Registra cuándo fue procesado el último archivo
- **mapeo_id_grupos**: Mapea IDs a nombres normalizados y sucursales

## Formato de Entrada

El sistema espera archivos .txt con el formato de WhatsApp:

```
[24/10/25, 10:51:52] Uzziel: Grupo BIENVENIDOS 
ID 000094
Pago 12921
Ahorro 1293 
Sucursal Ixtapaluca
```

## Funcionalidades de la GUI

1. **Subir Pagos**: Selecciona o arrastra archivos .txt de pagos
2. **Procesar Pagos**: Extrae y guarda los registros en Excel
3. **Subir Confirmaciones**: Selecciona archivos con confirmaciones
4. **Procesar Confirmaciones**: Marca como confirmados y crea hoja separada
5. **Ver Excel**: Abre el archivo Pagos.xlsx
6. **Limpiar Registros**: Elimina todos los datos (con confirmación)

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
- Corte (Matutino/Vespertino)
- Confirmado (Sí/No)

## Hojas del Excel

- **Pagos**: Todos los pagos registrados
- **Pagos Confirmados**: Pagos que han sido confirmados
- **Meta** (oculta): Información interna del sistema

## Notas

- Los nombres se normalizan a mayúsculas
- Las sucursales se normalizan quitando acentos
- El sistema previene duplicados por timestamp
- Los nuevos IDs de grupos se pueden agregar manualmente al config.json

## Autor

Sistema desarrollado para gestión de pagos desde WhatsApp

