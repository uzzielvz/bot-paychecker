# INSTRUCCIONES DE USO - Sistema WhatsApp → Excel

## Inicio Rápido

### Interfaz Gráfica (Recomendado)
```bash
# Doble clic en: iniciar_gui.bat
# O ejecutar: python gui.py
```

### Línea de Comandos
```bash
# Ejecutar: python main.py
```

## Exportar Chat de WhatsApp

1. Abre WhatsApp → Chat/Grupo
2. ⋮ → Más → Exportar chat
3. "Sin archivos multimedia"
4. Guardar en carpeta `input/`

## Funcionalidades

### 🖥️ Interfaz Gráfica Minimalista
La interfaz es simple y directa:
- **Zona de Selección**: Haz clic para elegir archivos de WhatsApp
- **Procesar Archivos**: Procesa todos los archivos pendientes en input/
- **Generar Excel**: Exporta todo a `output/pagos.xlsx`
- **Log Simple**: Mensajes claros con íconos (ℹ ✓ ✗)
- **Footer con Stats**: Registros totales, suma de pagos y ahorros

### 📟 Línea de Comandos (Alternativa)
1. **Monitoreo Automático**: Detecta y procesa automáticamente
2. **Generar Excel Actualizado**: Exporta todos los datos
3. **Procesar Archivos Pendientes**: Procesa todo en `input/`

## Carpetas

```
input/      → Coloca archivos .txt de WhatsApp aquí
output/     → Excel generado (pagos.xlsx)
database/   → Base de datos SQLite
logs/       → Registros de actividad
processed/  → Archivos ya procesados (respaldo)
```

## Notas

- Puedes procesar el mismo archivo varias veces (evita duplicados automáticamente)
- Configuración de cortes horarios en `config.json`
- Revisa `logs/procesamiento.log` si hay errores

