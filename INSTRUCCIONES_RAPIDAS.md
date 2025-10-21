# 🚀 Inicio Rápido

## Iniciar la Aplicación

### Opción 1: Doble Click (Windows)
Haz doble clic en:
```
iniciar_gui.bat
```

### Opción 2: Línea de Comandos
```bash
python gui.py
```

## Usar el Sistema

### 1️⃣ Exportar Chat de WhatsApp
1. Abre WhatsApp en tu teléfono
2. Ve al chat del grupo con los pagos
3. Toca **⋮** (menú) → **Más** → **Exportar chat**
4. Selecciona **"Sin archivos multimedia"**
5. Envía el archivo a tu computadora

### 2️⃣ Seleccionar Archivos
La interfaz tiene una zona grande en la parte superior:
- **Haz clic en la zona blanca** para seleccionar archivos
- Selecciona uno o varios archivos `.txt` de WhatsApp
- Los archivos se procesarán automáticamente

### 3️⃣ Procesar Archivos
Dos formas:
- **Botón "Procesar Archivos"**: Procesa todo lo que esté en la carpeta `input/`
- **Seleccionar archivos**: Haz clic en la zona de selección (arriba)

### 4️⃣ Generar Excel
- Haz clic en **"Generar Excel"**
- El archivo se guarda en `output/pagos.xlsx`
- Verás la ruta completa en el log

## Vista de la Interfaz

La interfaz es simple y limpia:
- **Zona de Selección**: Área grande para elegir archivos
- **2 Botones**: "Procesar Archivos" y "Generar Excel"
- **Log de Mensajes**: Muestra lo que está pasando con íconos:
  - ℹ Información
  - ✓ Éxito
  - ✗ Error
- **Footer**: Muestra totales (Registros, Pagos, Ahorros)

## ❓ Problemas Comunes

### No detecta archivos
✅ Verifica que el archivo esté en `input/` y tenga extensión `.txt`

### Error al generar Excel
✅ Asegúrate de que haya datos en la base de datos primero

### El monitoreo no inicia
✅ Cierra la app y vuélvela a abrir

## 📊 Formato de Pagos

El sistema detecta estos formatos en WhatsApp:

**Formato 1** (Compacto):
```
Nueva Luz 000031/ Puebla/ pago $5,852/ ahorro/ $ 348
```

**Formato 2** (Multilínea):
```
Grupo CATALEYA
ID000080
PAGO 7606
AHORRO 644
```

## 🔧 Configuración Avanzada

Edita `config.json` para cambiar:
- Horarios de cortes (matutino, vespertino, tarde)
- Carpetas del sistema
- Intervalo de monitoreo

