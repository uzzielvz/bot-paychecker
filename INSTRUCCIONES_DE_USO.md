# 📋 INSTRUCCIONES DE USO - Sistema WhatsApp → Excel

## 🚀 Inicio Rápido

### 1. Activar el Entorno Virtual

Cada vez que quieras usar el sistema, primero activa el entorno virtual:

```bash
# Windows PowerShell
.\venv\Scripts\activate.ps1

# O si usas CMD
venv\Scripts\activate.bat
```

### 2. Ejecutar el Sistema

```bash
python main.py
```

---

## 📱 Cómo Exportar Chat de WhatsApp

### Desde tu Teléfono:

1. Abre WhatsApp y ve al grupo/chat que quieres exportar
2. Toca los **3 puntos (⋮)** en la esquina superior derecha
3. Selecciona **"Más" → "Exportar chat"**
4. Elige **"Sin archivos multimedia"** (importante)
5. Envía el archivo a tu computadora (por email, Drive, etc.)
6. Descarga el archivo .txt a tu computadora

---

## 🎯 Opciones del Sistema

### Opción 1: Monitoreo Continuo (Recomendado)

**Cuándo usar:** Para procesar mensajes automáticamente durante todo el día.

**Pasos:**
1. Ejecuta `python main.py`
2. Selecciona opción `1`
3. El sistema queda vigilando la carpeta `input/`
4. Cuando copies un archivo .txt nuevo, lo procesará automáticamente
5. Presiona `Ctrl+C` para detener

**Ventajas:**
- Automático, solo copia el archivo y espera
- Procesa en 1-2 minutos después de copiar el archivo
- Ideal para uso durante todo el día

---

### Opción 2: Procesar Archivo Específico

**Cuándo usar:** Para procesar un archivo puntual, una sola vez.

**Pasos:**
1. Copia el archivo .txt a la carpeta `input/`
2. Ejecuta `python main.py`
3. Selecciona opción `2`
4. Elige el archivo de la lista
5. ¡Listo!

**Ventajas:**
- Control total sobre qué y cuándo procesar
- Más rápido (no espera, procesa de inmediato)

---

### Opción 3: Ver Estadísticas

Muestra un resumen completo de todos los pagos procesados:
- Total de registros
- Sumas de pagos y ahorros
- Resumen por corte horario
- Top 10 grupos
- Totales por sucursal

---

### Opción 4: Actualizar Excel

Regenera el archivo Excel desde la base de datos con todos los pagos históricos.

---

### Opción 5: Procesar Todo input/

Procesa todos los archivos .txt que estén en la carpeta `input/` de una sola vez.

---

## 📂 Estructura de Carpetas

```
agent-wppexcel/
├── input/              👈 COLOCA AQUÍ los archivos .txt de WhatsApp
├── output/             👈 AQUÍ SE GENERA pagos.xlsx
│   └── pagos.xlsx
├── database/           👈 Base de datos con historial
│   └── pagos.db
├── logs/               👈 Registros de actividad
│   └── procesamiento.log
└── processed/          👈 Archivos ya procesados (respaldo)
```

---

## ✅ Flujo de Trabajo Recomendado

### Para uso diario:

**1. Por la mañana:**
```bash
# Activar entorno
.\venv\Scripts\activate.ps1

# Iniciar monitoreo
python main.py
# Seleccionar opción 1
```

**2. Durante el día:**
- Exporta el chat de WhatsApp cada vez que necesites
- Copia el archivo .txt a la carpeta `input/`
- El sistema lo procesa automáticamente

**3. Al final del día:**
- Presiona `Ctrl+C` para detener el monitoreo
- Abre `output/pagos.xlsx` para ver todos los pagos del día

---

### Para uso esporádico:

**Cuando necesites procesar:**
```bash
# 1. Activa el entorno
.\venv\Scripts\activate.ps1

# 2. Copia el archivo .txt a input/

# 3. Ejecuta y selecciona opción 2
python main.py
```

---

## 🎨 Formatos de Mensajes Reconocidos

El sistema reconoce estos formatos:

### Formato Compacto (una línea):
```
Nueva Luz 000031/ Puebla/ pago $5,852/ ahorro/ $ 348
```

### Formato Multilínea:
```
Grupo CATALEYA
ID000080
PAGO 7606
AHORRO 644
```

**Variaciones aceptadas:**
- Con o sin símbolo `$`
- Con o sin comas en números
- Mayúsculas, minúsculas o mixto
- Espacios variables

---

## 🔄 Prevención de Duplicados

**¿Puedo procesar el mismo archivo varias veces?**
✅ **SÍ**, sin problema. El sistema detecta automáticamente mensajes duplicados.

**¿Puedo exportar el chat completo cada vez?**
✅ **SÍ**, recomendado. Solo procesará los mensajes nuevos.

**Ejemplo:**
- Lunes 9:00 AM: Exportas chat (10 pagos) → Procesa los 10
- Lunes 3:00 PM: Exportas el mismo chat (13 pagos) → Procesa solo los 3 nuevos
- Martes 10:00 AM: Exportas completo (20 pagos) → Procesa solo los 7 nuevos

---

## 📊 Archivos Generados

### 1. Excel (`output/pagos.xlsx`)

**Hoja "Pagos":**
- Todos los pagos con detalles completos
- Filtros automáticos en encabezados
- Formato de moneda
- Ordenado por fecha descendente

**Hoja "Resumen":**
- Total de registros y sumas
- Resumen por corte horario
- Totales por grupo y sucursal

### 2. Base de Datos (`database/pagos.db`)

Base de datos SQLite con todo el historial. Puedes consultarla con herramientas como DB Browser.

### 3. Logs (`logs/procesamiento.log`)

Registro detallado de todas las operaciones:
- Qué archivos se procesaron
- Cuántos pagos se encontraron
- Errores si los hubo

---

## ⚙️ Configuración de Cortes Horarios

Para cambiar los horarios de los cortes, edita `config.json`:

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

**Puedes:**
- Cambiar los horarios
- Cambiar los nombres
- Agregar más cortes (ej: Corte Nocturno)

---

## ❗ Solución de Problemas

### "No module named 'openpyxl'"

**Solución:**
```bash
.\venv\Scripts\activate.ps1
pip install -r requirements.txt
```

### "No se encontraron pagos"

**Verifica:**
- ¿El mensaje tiene el formato correcto?
- ¿Exportaste con "Sin archivos multimedia"?
- Revisa `logs/procesamiento.log` para más detalles

### El monitoreo no detecta archivos

**Verifica:**
- ¿El archivo tiene extensión `.txt`?
- ¿Está en la carpeta `input/` correcta?
- Espera 1-2 minutos (intervalo de revisión)

### Error al abrir Excel

Si el Excel está abierto cuando el sistema intenta actualizarlo, ciérralo y vuelve a procesar.

---

## 💡 Consejos

1. **Exporta el chat completo** cada vez (no te preocupes por duplicados)
2. **Usa nombres descriptivos** para los archivos:
   - ✅ `pagos_16oct_mañana.txt`
   - ✅ `grupo_principal_octubre.txt`
   - ❌ `chat.txt`

3. **Revisa los logs** si algo no funciona como esperabas

4. **Haz respaldo** del Excel y base de datos periódicamente

5. **No elimines** la carpeta `processed/` (contiene respaldos)

---

## 📞 Ayuda Adicional

Si algo no funciona:

1. Revisa `logs/procesamiento.log`
2. Verifica que el formato del mensaje sea correcto
3. Prueba con la opción 2 (procesar manual) primero
4. Asegúrate de haber activado el entorno virtual

---

**¡Listo! Ya puedes usar el sistema para procesar tus pagos de WhatsApp automáticamente.** 🎉

