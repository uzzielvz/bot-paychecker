#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Sistema de Extracción de Pagos de WhatsApp a Excel
Procesa archivos de chat exportados de WhatsApp y extrae información de pagos
"""

import os
import sys
from pathlib import Path
from monitor import Monitor, FileProcessor
from database_manager import DatabaseManager
from excel_manager import ExcelManager


def mostrar_menu():
    """Muestra el menú principal"""
    print("\n" + "="*60)
    print(" SISTEMA DE EXTRACCIÓN DE PAGOS WHATSAPP → EXCEL")
    print("="*60)
    print("\nOpciones:")
    print("  1. Iniciar monitoreo automático")
    print("  2. Generar Excel actualizado")
    print("  3. Procesar todos los archivos pendientes")
    print("  0. Salir")
    print("-"*60)


def mostrar_estadisticas():
    """Muestra estadísticas de la base de datos"""
    try:
        db = DatabaseManager()
        stats = db.obtener_estadisticas()
        
        print("\n" + "="*60)
        print(" ESTADÍSTICAS DE LA BASE DE DATOS")
        print("="*60)
        
        print(f"\nTotal de registros: {stats['total_pagos']}")
        print(f"Suma total de pagos: ${stats['suma_total_pago']:,.2f}")
        print(f"Suma total de ahorros: ${stats['suma_total_ahorro']:,.2f}")
        
        print("\n" + "-"*60)
        print("RESUMEN POR CORTE HORARIO:")
        print("-"*60)
        print(f"{'Corte':<20} {'Cantidad':>10} {'Total Pago':>15} {'Total Ahorro':>15}")
        print("-"*60)
        
        for corte, cantidad, pago, ahorro in stats['por_corte']:
            print(f"{corte:<20} {cantidad:>10} ${pago:>14,.2f} ${ahorro:>14,.2f}")
        
        print("\n" + "-"*60)
        print("TOP 10 GRUPOS:")
        print("-"*60)
        print(f"{'Grupo':<25} {'Cantidad':>10} {'Total Pago':>15}")
        print("-"*60)
        
        for grupo, cantidad, pago, ahorro in stats['por_grupo'][:10]:
            print(f"{grupo:<25} {cantidad:>10} ${pago:>14,.2f}")
        
        print("\n" + "-"*60)
        print("POR SUCURSAL:")
        print("-"*60)
        print(f"{'Sucursal':<20} {'Cantidad':>10} {'Total Pago':>15}")
        print("-"*60)
        
        for sucursal, cantidad, pago, ahorro in stats['por_sucursal']:
            print(f"{sucursal:<20} {cantidad:>10} ${pago:>14,.2f}")
        
        print("="*60)
        
    except Exception as e:
        print(f"\nError al obtener estadísticas: {e}")


def procesar_archivo_especifico():
    """Procesa un archivo específico seleccionado por el usuario"""
    print("\n" + "-"*60)
    print("Archivos disponibles en input/:")
    print("-"*60)
    
    input_dir = "input/"
    archivos = list(Path(input_dir).glob('*.txt'))
    
    if not archivos:
        print("No hay archivos .txt en la carpeta input/")
        print("\nColoca un archivo de chat exportado de WhatsApp en:")
        print(f"  {os.path.abspath(input_dir)}")
        return
    
    for i, archivo in enumerate(archivos, 1):
        print(f"  {i}. {archivo.name}")
    
    print("-"*60)
    
    try:
        opcion = input("\nSelecciona el número del archivo (0 para cancelar): ").strip()
        
        if opcion == '0':
            return
        
        indice = int(opcion) - 1
        
        if 0 <= indice < len(archivos):
            archivo_seleccionado = str(archivos[indice])
            print(f"\nProcesando: {archivos[indice].name}")
            
            processor = FileProcessor()
            resultado = processor.procesar_archivo(archivo_seleccionado)
            
            print("\n" + "="*60)
            print("RESULTADO DEL PROCESAMIENTO")
            print("="*60)
            print(f"Archivo: {resultado['archivo']}")
            
            if 'error' in resultado:
                print(f"ERROR: {resultado['error']}")
            else:
                print(f"Pagos encontrados: {resultado['pagos_encontrados']}")
                print(f"Nuevos insertados: {resultado['insertados']}")
                print(f"Duplicados (ignorados): {resultado['duplicados']}")
                print(f"Errores: {resultado['errores']}")
            
            print("="*60)
        else:
            print("Opción inválida")
    
    except ValueError:
        print("Por favor ingresa un número válido")
    except Exception as e:
        print(f"Error: {e}")


def generar_excel():
    """Genera el Excel desde la base de datos"""
    try:
        print("\nGenerando Excel desde la base de datos...")
        
        db = DatabaseManager()
        excel_manager = ExcelManager()
        
        pagos = db.obtener_todos_pagos()
        
        if not pagos:
            print("\nNo hay datos en la base de datos para generar Excel")
            return
        
        excel_manager.generar_excel(pagos)
        
        print("\n" + "="*60)
        print("EXCEL GENERADO EXITOSAMENTE")
        print("="*60)
        print(f"Total de registros: {len(pagos)}")
        print(f"Ubicación: {os.path.abspath(excel_manager.excel_path)}")
        print("="*60)
        
    except Exception as e:
        print(f"\nError al generar Excel: {e}")


def procesar_todos_archivos():
    """Procesa todos los archivos en la carpeta input/"""
    print("\nProcesando todos los archivos en input/...")
    
    processor = FileProcessor()
    processor.procesar_archivos_existentes()
    
    print("\nProcesamiento completado")


def main():
    """Función principal"""
    print("\nSistema de Extracción de Pagos WhatsApp → Excel")
    
    # Verificar que las carpetas existan
    carpetas = ['input', 'output', 'database', 'logs', 'processed']
    for carpeta in carpetas:
        os.makedirs(carpeta, exist_ok=True)
    
    while True:
        mostrar_menu()
        
        try:
            opcion = input("\nSelecciona una opción: ").strip()
            
            if opcion == '1':
                # Monitoreo continuo
                monitor = Monitor()
                monitor.iniciar_monitoreo()
            
            elif opcion == '2':
                # Generar Excel
                generar_excel()
            
            elif opcion == '3':
                # Procesar todos los archivos
                procesar_todos_archivos()
            
            elif opcion == '0':
                print("\nSistema cerrado correctamente")
                sys.exit(0)
            
            else:
                print("\nOpción no válida. Por favor selecciona 0-3.")
        
        except KeyboardInterrupt:
            print("\n\nSistema cerrado correctamente")
            sys.exit(0)
        except Exception as e:
            print(f"\nError: {e}")
            input("\nPresiona Enter para continuar...")


if __name__ == "__main__":
    main()

