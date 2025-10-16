import os
import time
import json
import shutil
import logging
from datetime import datetime
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from extractor import WhatsAppExtractor
from database_manager import DatabaseManager
from excel_manager import ExcelManager


class WhatsAppFileHandler(FileSystemEventHandler):
    """Maneja eventos de archivos nuevos en la carpeta input/"""
    
    def __init__(self, processor):
        self.processor = processor
        self.processed_files = set()
    
    def on_created(self, event):
        """Se ejecuta cuando se detecta un archivo nuevo"""
        if event.is_directory:
            return
        
        if event.src_path.endswith('.txt'):
            # Esperar un momento para asegurar que el archivo esté completamente escrito
            time.sleep(1)
            self.processor.procesar_archivo(event.src_path)
    
    def on_modified(self, event):
        """Se ejecuta cuando un archivo es modificado"""
        if event.is_directory:
            return
        
        if event.src_path.endswith('.txt') and event.src_path not in self.processed_files:
            time.sleep(1)
            self.processor.procesar_archivo(event.src_path)
            self.processed_files.add(event.src_path)


class FileProcessor:
    """Procesa archivos de WhatsApp y coordina extracción, BD y Excel"""
    
    def __init__(self, config_path: str = "config.json"):
        """Inicializa el procesador"""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.extractor = WhatsAppExtractor(config_path)
        self.db_manager = DatabaseManager(config_path)
        self.excel_manager = ExcelManager(config_path)
        
        # Configurar logging
        log_dir = self.config['rutas']['logs']
        os.makedirs(log_dir, exist_ok=True)
        
        log_path = os.path.join(log_dir, 'procesamiento.log')
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_path, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger(__name__)
    
    def procesar_archivo(self, ruta_archivo: str) -> dict:
        """Procesa un archivo de chat de WhatsApp completo"""
        nombre_archivo = os.path.basename(ruta_archivo)
        self.logger.info(f"Iniciando procesamiento de: {nombre_archivo}")
        
        try:
            # Extraer pagos del archivo
            pagos = self.extractor.procesar_archivo(ruta_archivo)
            self.logger.info(f"Pagos encontrados: {len(pagos)}")
            
            if not pagos:
                self.logger.warning(f"No se encontraron pagos en {nombre_archivo}")
                return {
                    'archivo': nombre_archivo,
                    'pagos_encontrados': 0,
                    'insertados': 0,
                    'duplicados': 0,
                    'errores': 0
                }
            
            # Insertar en base de datos
            stats = self.db_manager.insertar_pagos_lote(pagos, nombre_archivo)
            self.logger.info(
                f"Resultados: {stats['insertados']} insertados, "
                f"{stats['duplicados']} duplicados, {stats['errores']} errores"
            )
            
            # Actualizar Excel con todos los pagos de la BD
            todos_pagos = self.db_manager.obtener_todos_pagos()
            self.excel_manager.generar_excel(todos_pagos)
            self.logger.info("Excel actualizado correctamente")
            
            # Mover archivo a processed/
            self._mover_a_procesados(ruta_archivo)
            
            return {
                'archivo': nombre_archivo,
                'pagos_encontrados': len(pagos),
                'insertados': stats['insertados'],
                'duplicados': stats['duplicados'],
                'errores': stats['errores']
            }
            
        except Exception as e:
            self.logger.error(f"Error procesando {nombre_archivo}: {str(e)}", exc_info=True)
            return {
                'archivo': nombre_archivo,
                'error': str(e)
            }
    
    def _mover_a_procesados(self, ruta_archivo: str):
        """Mueve un archivo procesado a la carpeta processed/ con timestamp"""
        processed_dir = self.config['rutas']['processed']
        os.makedirs(processed_dir, exist_ok=True)
        
        nombre_archivo = os.path.basename(ruta_archivo)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nombre_base, extension = os.path.splitext(nombre_archivo)
        nuevo_nombre = f"{nombre_base}_{timestamp}{extension}"
        
        destino = os.path.join(processed_dir, nuevo_nombre)
        
        try:
            shutil.move(ruta_archivo, destino)
            self.logger.info(f"Archivo movido a: {destino}")
        except Exception as e:
            self.logger.error(f"Error moviendo archivo: {e}")
    
    def procesar_archivos_existentes(self):
        """Procesa todos los archivos .txt que ya están en input/"""
        input_dir = self.config['rutas']['input']
        archivos_txt = Path(input_dir).glob('*.txt')
        
        for archivo in archivos_txt:
            self.procesar_archivo(str(archivo))


class Monitor:
    """Monitor de carpeta que detecta archivos nuevos"""
    
    def __init__(self, config_path: str = "config.json"):
        """Inicializa el monitor"""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.processor = FileProcessor(config_path)
        self.input_dir = self.config['rutas']['input']
        self.intervalo = self.config.get('intervalo_monitoreo_minutos', 2) * 60
        
        # Asegurar que existe la carpeta input
        os.makedirs(self.input_dir, exist_ok=True)
    
    def iniciar_monitoreo(self):
        """Inicia el monitoreo continuo de la carpeta input/"""
        print(f"\n{'='*60}")
        print(f"MONITOR DE ARCHIVOS INICIADO")
        print(f"{'='*60}")
        print(f"Carpeta monitoreada: {os.path.abspath(self.input_dir)}")
        print(f"Intervalo de revisión: {self.intervalo // 60} minutos")
        print(f"Presiona Ctrl+C para detener el monitoreo")
        print(f"{'='*60}\n")
        
        # Procesar archivos existentes primero
        print("Buscando archivos existentes...")
        self.processor.procesar_archivos_existentes()
        
        # Configurar watchdog
        event_handler = WhatsAppFileHandler(self.processor)
        observer = Observer()
        observer.schedule(event_handler, self.input_dir, recursive=False)
        observer.start()
        
        try:
            while True:
                time.sleep(self.intervalo)
                # Re-escanear por si acaso watchdog no detectó algo
                self.processor.procesar_archivos_existentes()
        except KeyboardInterrupt:
            print("\n\nDeteniendo monitor...")
            observer.stop()
        
        observer.join()
        print("Monitor detenido correctamente")


if __name__ == "__main__":
    # Prueba básica
    monitor = Monitor()
    print("Monitor inicializado correctamente")
    print(f"Carpeta de entrada: {monitor.input_dir}")

