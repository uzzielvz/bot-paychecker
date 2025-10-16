# Sistema de Extracción de Pagos WhatsApp → Excel

Sistema automatizado que procesa archivos de chat de WhatsApp exportados, extrae información de pagos de grupos y la almacena en Excel y base de datos SQLite.

## Características

- ✅ Extracción automática de pagos desde chats de WhatsApp
- ✅ Soporte para múltiples formatos de mensajes
- ✅ Prevención de duplicados (puedes procesar el mismo archivo múltiples veces)
- ✅ Categorización por cortes horarios (Matutino, Vespertino, Tarde)
- ✅ Almacenamiento en base de datos SQLite
- ✅ Generación automática de Excel con formato profesional
- ✅ Monitoreo continuo de carpeta input/
- ✅ Estadísticas y resúmenes detallados

## Requisitos

- Python 3.7 o superior
- Windows, Linux o macOS

## Instalación

### 1. Clonar o descargar el proyecto

```bash
cd agent-wppexcel
```

### 2. Crear entorno virtual

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/Mac:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

## Uso

### Opción 1: Monitoreo Automático (Recomendado)

1. Activa el entorno virtual
2. Ejecuta el programa:
   ```bash
   python main.py
   ```
3. Selecciona opción `1` (Iniciar monitoreo continuo)
4. El sistema quedará vigilando la carpeta `input/`
5. Exporta un chat desde WhatsApp y colócalo en la carpeta `input/`
6. El sistema lo detectará y procesará automáticamente

### Opción 2: Procesamiento Manual

1. Coloca un archivo .txt de WhatsApp en la carpeta `input/`
2. Ejecuta:
   ```bash
   python main.py
   ```
3. Selecciona opción `2` (Procesar archivo específico)
4. Elige el archivo a procesar

## Cómo exportar chat de WhatsApp

### Desde tu teléfono:

1. Abre el chat en WhatsApp
2. Toca los 3 puntos (⋮) → Más → Exportar chat
3. Selecciona "Sin archivos multimedia"
4. Envía el archivo a tu computadora (email, Drive, etc.)
5. Coloca el archivo en la carpeta `input/`

## Formatos de Mensajes Soportados

El sistema reconoce estos formatos de pagos:

**Formato compacto:**
```
Nueva Luz 000031/ Puebla/ pago $5,852/ ahorro/ $ 348
```

**Formato multilínea:**
```
Grupo CATALEYA
ID000080
PAGO 7606
AHORRO 644
```

## Estructura del Proyecto

```
agent-wppexcel/
├── input/              # Coloca aquí los .txt exportados
├── output/             # Excel generado: pagos.xlsx
├── database/           # Base de datos: pagos.db
├── logs/               # Registros de procesamiento
├── processed/          # Archivos ya procesados (respaldo)
├── config.json         # Configuración (cortes horarios)
├── extractor.py        # Lógica de extracción
├── database_manager.py # Gestión de base de datos
├── excel_manager.py    # Generación de Excel
├── monitor.py          # Monitoreo de carpeta
└── main.py             # Script principal
```

## Configuración de Cortes Horarios

Edita `config.json` para personalizar los cortes:

```json
{
  "cortes_horarios": [
    {
      "nombre": "Corte Matutino",
      "hora_inicio": "09:00",
      "hora_fin": "12:00"
    },
    {
      "nombre": "Corte Vespertino",
      "hora_inicio": "12:01",
      "hora_fin": "15:00"
    },
    {
      "nombre": "Corte Tarde",
      "hora_inicio": "15:01",
      "hora_fin": "23:59"
    }
  ]
}
```

## Prevención de Duplicados

El sistema usa hashes SHA256 para identificar mensajes únicos. Puedes:

- ✅ Procesar el mismo archivo múltiples veces
- ✅ Exportar el chat completo cada vez (solo procesa los nuevos)
- ✅ No preocuparte por duplicados

## Archivos de Salida

- **`output/pagos.xlsx`** - Excel con todos los pagos procesados
  - Hoja "Pagos": Datos completos
  - Hoja "Resumen": Estadísticas por corte y grupo
  
- **`database/pagos.db`** - Base de datos SQLite con historial completo

- **`logs/procesamiento.log`** - Registro de todas las operaciones

## Solución de Problemas

### Error: "No module named 'openpyxl'"
```bash
pip install -r requirements.txt
```

### El sistema no detecta archivos nuevos
- Verifica que el archivo tenga extensión .txt
- Asegúrate de que esté en la carpeta `input/`
- Espera 1-2 minutos (intervalo de monitoreo)

### No se extraen pagos de un mensaje
- Verifica que el formato sea correcto
- Revisa el archivo `logs/procesamiento.log` para más detalles

## Consumo de Recursos

- **Memoria RAM:** ~35-40 MB
- **Almacenamiento (1 año):** ~100-200 MB
- **CPU:** Mínimo, solo al procesar archivos

## Autor

Desarrollado por Uzziel Valdez

## Licencia

MIT License

