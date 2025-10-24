#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Interfaz Gráfica Minimalista - Sistema de Extracción de Pagos WhatsApp → Excel
"""

import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import threading
import os
import shutil
from pathlib import Path
from datetime import datetime
import queue
from tkinterdnd2 import DND_FILES, TkinterDnD

from monitor import FileProcessor
from database_manager import DatabaseManager
from excel_manager import ExcelManager


class MinimalButton(tk.Button):
    """Botón minimalista con hover sutil"""
    def __init__(self, parent, **kwargs):
        super().__init__(
            parent,
            relief=tk.FLAT,
            cursor="hand2",
            font=("Segoe UI", 11),
            bd=0,
            highlightthickness=0,
            **kwargs
        )
        self.default_bg = kwargs.get('bg', '#666666')
        self.hover_bg = '#555555'
        self.bind("<Enter>", lambda e: self.config(bg=self.hover_bg))
        self.bind("<Leave>", lambda e: self.config(bg=self.default_bg))


class PagosExtractorGUI:
    """Interfaz gráfica minimalista del sistema"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Extracción de Pagos WhatsApp")
        self.root.geometry("700x600")
        self.root.minsize(700, 600)
        
        # Colores minimalistas
        self.colors = {
            'bg': '#F5F5F5',
            'fg': '#333333',
            'button': '#666666',
            'accent': '#5C6B7A',
            'white': '#FFFFFF',
            'border': '#DDDDDD',
            'text_bg': '#FAFAFA'
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Variables
        self.log_queue = queue.Queue()
        self.is_processing = False
        
        # Inicializar componentes
        self.processor = FileProcessor()
        self.db_manager = DatabaseManager()
        self.excel_manager = ExcelManager()
        
        # Verificar carpetas
        self._verificar_carpetas()
        
        # Configurar UI
        self._crear_interfaz()
        
        # Actualizar log periódicamente
        self._actualizar_log()
        
        # Cargar estadísticas iniciales
        self._actualizar_estadisticas()
    
    def _verificar_carpetas(self):
        """Verifica y crea las carpetas necesarias"""
        carpetas = ['input', 'output', 'database', 'logs', 'processed']
        for carpeta in carpetas:
            os.makedirs(carpeta, exist_ok=True)
    
    def _crear_interfaz(self):
        """Crea la interfaz minimalista"""
        # Header simple
        header_frame = tk.Frame(self.root, bg=self.colors['white'], height=60)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="Extracción de Pagos WhatsApp",
            font=("Segoe UI", 16),
            bg=self.colors['white'],
            fg=self.colors['fg']
        )
        title_label.pack(pady=18)
        
        # Contenedor principal
        main_container = tk.Frame(self.root, bg=self.colors['bg'])
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Zona de drop
        self._crear_zona_drop(main_container)
        
        # Botones principales
        self._crear_botones(main_container)
        
        # Área de mensajes
        self._crear_area_mensajes(main_container)
        
        # Footer con estadísticas
        self._crear_footer()
    
    def _crear_zona_drop(self, parent):
        """Crea la zona de selección de archivos con soporte para drag and drop"""
        drop_frame = tk.Frame(
            parent,
            bg=self.colors['white'],
            relief=tk.FLAT,
            bd=2,
            highlightbackground=self.colors['border'],
            highlightthickness=2
        )
        drop_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Instrucciones
        info_label = tk.Label(
            drop_frame,
            text="Arrastra archivos aquí o haz clic para seleccionar",
            font=("Segoe UI", 11),
            bg=self.colors['white'],
            fg=self.colors['fg'],
            cursor="hand2",
            justify=tk.CENTER
        )
        info_label.pack(pady=40, padx=20)
        
        # Hacer que sea clickeable
        info_label.bind("<Button-1>", lambda e: self._seleccionar_archivos())
        drop_frame.bind("<Button-1>", lambda e: self._seleccionar_archivos())
        
        # Configurar drag and drop
        drop_frame.drop_target_register(DND_FILES)
        drop_frame.dnd_bind('<<Drop>>', self._on_drop)
        drop_frame.dnd_bind('<<DragEnter>>', self._on_drag_enter)
        drop_frame.dnd_bind('<<DragLeave>>', self._on_drag_leave)
        
        # Efecto hover
        def on_enter(e):
            drop_frame.config(highlightbackground=self.colors['accent'])
            info_label.config(fg=self.colors['accent'])
        
        def on_leave(e):
            drop_frame.config(highlightbackground=self.colors['border'])
            info_label.config(fg=self.colors['fg'])
        
        drop_frame.bind("<Enter>", on_enter)
        drop_frame.bind("<Leave>", on_leave)
        info_label.bind("<Enter>", on_enter)
        info_label.bind("<Leave>", on_leave)
        
        self.drop_frame = drop_frame
        self.info_label = info_label
    
    def _on_drag_enter(self, event):
        """Maneja el evento cuando se arrastra un archivo sobre la zona"""
        self.drop_frame.config(highlightbackground=self.colors['accent'])
        self.info_label.config(fg=self.colors['accent'])
        self.info_label.config(text="¡Suelta aquí para copiar archivos!")
    
    def _on_drag_leave(self, event):
        """Maneja el evento cuando se deja de arrastrar sobre la zona"""
        self.drop_frame.config(highlightbackground=self.colors['border'])
        self.info_label.config(fg=self.colors['fg'])
        self.info_label.config(text="Arrastra archivos aquí o haz clic para seleccionar")
    
    def _on_drop(self, event):
        """Maneja el evento cuando se sueltan archivos en la zona"""
        # Restaurar apariencia normal
        self.drop_frame.config(highlightbackground=self.colors['border'])
        self.info_label.config(fg=self.colors['fg'])
        self.info_label.config(text="Arrastra archivos aquí o haz clic para seleccionar")
        
        # Obtener lista de archivos
        archivos = self.root.tk.splitlist(event.data)
        
        # Filtrar solo archivos .txt
        archivos_txt = [archivo for archivo in archivos if archivo.lower().endswith('.txt')]
        
        if not archivos_txt:
            messagebox.showwarning("Archivos no válidos", "Solo se pueden procesar archivos .txt")
            return
        
        # Copiar archivos a input/ y procesarlos
        self._copiar_y_procesar_archivos(archivos_txt)
    
    def _copiar_y_procesar_archivos(self, archivos_origen):
        """Copia archivos a input/ y los procesa"""
        if self.is_processing:
            messagebox.showwarning("Procesando", "Ya hay un procesamiento en curso")
            return
        
        def copiar_y_procesar():
            self.is_processing = True
            try:
                self._log(f"ℹ Copiando {len(archivos_origen)} archivo(s)...")
                
                archivos_copiados = []
                for archivo_origen in archivos_origen:
                    nombre = os.path.basename(archivo_origen)
                    destino = os.path.join("input", nombre)
                    
                    # Si el archivo ya existe, agregar timestamp
                    if os.path.exists(destino):
                        nombre_base, extension = os.path.splitext(nombre)
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        nombre = f"{nombre_base}_{timestamp}{extension}"
                        destino = os.path.join("input", nombre)
                    
                    # Copiar archivo
                    shutil.copy2(archivo_origen, destino)
                    archivos_copiados.append(destino)
                    self._log(f"✓ Copiado: {nombre}")
                
                # Procesar archivos copiados
                self._log(f"ℹ Procesando {len(archivos_copiados)} archivo(s)...")
                
                total_insertados = 0
                total_duplicados = 0
                
                for archivo in archivos_copiados:
                    nombre = os.path.basename(archivo)
                    self._log(f"ℹ {nombre}")
                    
                    resultado = self.processor.procesar_archivo(archivo)
                    
                    if 'error' in resultado:
                        error_msg = resultado['error']
                        if "Permission denied" in str(error_msg) and "pagos.xlsx" in str(error_msg):
                            self._log(f"  ⚠ Excel abierto - se creó respaldo")
                        else:
                            self._log(f"  ✗ Error: {error_msg}")
                    else:
                        total_insertados += resultado['insertados']
                        total_duplicados += resultado['duplicados']
                        self._log(f"  ✓ {resultado['insertados']} nuevos, {resultado['duplicados']} duplicados")
                
                self._actualizar_estadisticas()
                self._log(f"✓ Completado: {total_insertados} registros nuevos")
                
                # Verificar si se creó un respaldo
                respaldo_creado = any("Permission denied" in str(resultado.get('error', '')) and "pagos.xlsx" in str(resultado.get('error', '')) 
                                    for archivo in archivos_copiados 
                                    for resultado in [self.processor.procesar_archivo(archivo)] 
                                    if 'error' in resultado)
                
                mensaje = f"Archivos copiados y procesados\n\nNuevos: {total_insertados}\nDuplicados: {total_duplicados}"
                if respaldo_creado:
                    mensaje += "\n\n⚠ Se creó un Excel de respaldo porque el original está abierto"
                
                messagebox.showinfo("Éxito", mensaje)
                
            except Exception as e:
                self._log(f"✗ Error: {e}")
                messagebox.showerror("Error", f"Error al copiar/procesar archivos:\n{e}")
            finally:
                self.is_processing = False
        
        threading.Thread(target=copiar_y_procesar, daemon=True).start()
    
    def _crear_botones(self, parent):
        """Crea los botones principales"""
        button_frame = tk.Frame(parent, bg=self.colors['bg'])
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Botón Procesar
        btn_procesar = MinimalButton(
            button_frame,
            text="Procesar Archivos",
            bg=self.colors['button'],
            fg=self.colors['white'],
            width=15,
            height=2,
            command=self._procesar_archivos
        )
        btn_procesar.pack(side=tk.LEFT, expand=True, padx=(0, 5))
        
        # Botón Generar Excel
        btn_excel = MinimalButton(
            button_frame,
            text="Generar Excel",
            bg=self.colors['accent'],
            fg=self.colors['white'],
            width=15,
            height=2,
            command=self._generar_excel
        )
        btn_excel.pack(side=tk.LEFT, expand=True, padx=(5, 5))
        
        # Botón Abrir Excel
        btn_abrir = MinimalButton(
            button_frame,
            text="Abrir Excel",
            bg=self.colors['button'],
            fg=self.colors['white'],
            width=15,
            height=2,
            command=self._abrir_excel
        )
        btn_abrir.pack(side=tk.LEFT, expand=True, padx=(5, 0))
    
    def _crear_area_mensajes(self, parent):
        """Crea el área de mensajes simple"""
        msg_label = tk.Label(
            parent,
            text="Mensajes",
            font=("Segoe UI", 10),
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            anchor=tk.W
        )
        msg_label.pack(fill=tk.X, pady=(0, 5))
        
        # Área de texto simple
        text_frame = tk.Frame(parent, bg=self.colors['border'])
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(
            text_frame,
            wrap=tk.WORD,
            font=("Segoe UI", 9),
            bg=self.colors['text_bg'],
            fg=self.colors['fg'],
            insertbackground=self.colors['fg'],
            state=tk.DISABLED,
            relief=tk.FLAT,
            bd=0,
            padx=10,
            pady=10
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        self._log("ℹ Sistema listo")
    
    def _crear_footer(self):
        """Crea el footer con estadísticas mínimas"""
        footer_frame = tk.Frame(self.root, bg=self.colors['white'], height=50)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        # Contenedor para stats
        stats_container = tk.Frame(footer_frame, bg=self.colors['white'])
        stats_container.pack(expand=True)
        
        # Total Registros
        self.stat_registros = tk.Label(
            stats_container,
            text="Registros: 0",
            font=("Segoe UI", 9),
            bg=self.colors['white'],
            fg=self.colors['fg']
        )
        self.stat_registros.pack(side=tk.LEFT, padx=20)
        
        # Total Pagos
        self.stat_pagos = tk.Label(
            stats_container,
            text="Pagos: $0.00",
            font=("Segoe UI", 9),
            bg=self.colors['white'],
            fg=self.colors['fg']
        )
        self.stat_pagos.pack(side=tk.LEFT, padx=20)
        
        # Total Ahorros
        self.stat_ahorros = tk.Label(
            stats_container,
            text="Ahorros: $0.00",
            font=("Segoe UI", 9),
            bg=self.colors['white'],
            fg=self.colors['fg']
        )
        self.stat_ahorros.pack(side=tk.LEFT, padx=20)
    
    def _log(self, mensaje):
        """Agrega un mensaje al log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        mensaje_completo = f"[{timestamp}] {mensaje}\n"
        self.log_queue.put(mensaje_completo)
    
    def _actualizar_log(self):
        """Actualiza el área de log desde la cola"""
        try:
            while True:
                mensaje = self.log_queue.get_nowait()
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, mensaje)
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self._actualizar_log)
    
    def _actualizar_estadisticas(self):
        """Actualiza las estadísticas del footer"""
        try:
            stats = self.db_manager.obtener_estadisticas()
            
            self.stat_registros.config(text=f"Registros: {stats['total_pagos']:,}")
            self.stat_pagos.config(text=f"Pagos: ${stats['suma_total_pago']:,.2f}")
            self.stat_ahorros.config(text=f"Ahorros: ${stats['suma_total_ahorro']:,.2f}")
            
        except Exception as e:
            self._log(f"✗ Error al actualizar estadísticas: {e}")
    
    def _seleccionar_archivos(self):
        """Abre diálogo para seleccionar archivos"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos de WhatsApp",
            filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")],
            initialdir=os.path.abspath("input/")
        )
        
        if archivos:
            self._procesar_archivos_lista(list(archivos))
    
    def _procesar_archivos_lista(self, archivos):
        """Procesa una lista de archivos"""
        if self.is_processing:
            messagebox.showwarning("Procesando", "Ya hay un procesamiento en curso")
            return
        
        def procesar():
            self.is_processing = True
            try:
                self._log(f"ℹ Procesando {len(archivos)} archivo(s)...")
                
                total_insertados = 0
                total_duplicados = 0
                
                for archivo in archivos:
                    nombre = os.path.basename(archivo)
                    self._log(f"ℹ {nombre}")
                    
                    # Copiar a input si no está ahí
                    input_path = os.path.join("input", nombre)
                    if archivo != input_path and not os.path.exists(input_path):
                        import shutil
                        shutil.copy2(archivo, input_path)
                    
                    resultado = self.processor.procesar_archivo(input_path)
                    
                    if 'error' in resultado:
                        self._log(f"  ✗ Error: {resultado['error']}")
                    else:
                        total_insertados += resultado['insertados']
                        total_duplicados += resultado['duplicados']
                        self._log(f"  ✓ {resultado['insertados']} nuevos, {resultado['duplicados']} duplicados")
                
                self._actualizar_estadisticas()
                self._log(f"✓ Completado: {total_insertados} registros nuevos")
                
                messagebox.showinfo("Éxito", f"Procesamiento completado\n\nNuevos: {total_insertados}\nDuplicados: {total_duplicados}")
                
            except Exception as e:
                self._log(f"✗ Error: {e}")
                messagebox.showerror("Error", f"Error al procesar:\n{e}")
            finally:
                self.is_processing = False
        
        threading.Thread(target=procesar, daemon=True).start()
    
    def _procesar_archivos(self):
        """Procesa todos los archivos en input/"""
        if self.is_processing:
            messagebox.showwarning("Procesando", "Ya hay un procesamiento en curso")
            return
        
        def procesar():
            self.is_processing = True
            try:
                self._log("ℹ Buscando archivos en input/...")
                
                input_dir = "input/"
                archivos = list(Path(input_dir).glob('*.txt'))
                
                if not archivos:
                    self._log("✗ No hay archivos .txt en input/")
                    messagebox.showinfo("Sin archivos", "No hay archivos .txt en la carpeta input/")
                    self.is_processing = False
                    return
                
                self._log(f"ℹ Encontrados {len(archivos)} archivo(s)")
                
                total_insertados = 0
                total_duplicados = 0
                
                for archivo in archivos:
                    self._log(f"ℹ {archivo.name}")
                    resultado = self.processor.procesar_archivo(str(archivo))
                    
                    if 'error' in resultado:
                        self._log(f"  ✗ {resultado['error']}")
                    else:
                        total_insertados += resultado['insertados']
                        total_duplicados += resultado['duplicados']
                        self._log(f"  ✓ {resultado['insertados']} nuevos, {resultado['duplicados']} duplicados")
                
                self._actualizar_estadisticas()
                self._log(f"✓ Completado: {total_insertados} registros nuevos")
                
                messagebox.showinfo("Éxito", f"Procesamiento completado\n\nNuevos: {total_insertados}\nDuplicados: {total_duplicados}")
                
            except Exception as e:
                self._log(f"✗ Error: {e}")
                messagebox.showerror("Error", f"Error al procesar:\n{e}")
            finally:
                self.is_processing = False
        
        threading.Thread(target=procesar, daemon=True).start()
    
    def _generar_excel(self):
        """Genera el archivo Excel"""
        if self.is_processing:
            messagebox.showwarning("Procesando", "Espera a que termine el procesamiento")
            return
        
        def generar():
            self.is_processing = True
            try:
                self._log("ℹ Generando Excel...")
                
                pagos = self.db_manager.obtener_todos_pagos()
                
                if not pagos:
                    self._log("✗ No hay datos para generar Excel")
                    messagebox.showwarning("Sin datos", "No hay datos en la base de datos")
                    self.is_processing = False
                    return
                
                self.excel_manager.generar_excel(pagos)
                
                self._log(f"✓ Excel generado: {len(pagos)} registros")
                self._log(f"  📄 {os.path.abspath(self.excel_manager.excel_path)}")
                
                messagebox.showinfo(
                    "Éxito",
                    f"Excel generado exitosamente\n\n{len(pagos)} registros\n\n{os.path.abspath(self.excel_manager.excel_path)}"
                )
                
            except Exception as e:
                self._log(f"✗ Error: {e}")
                messagebox.showerror("Error", f"Error al generar Excel:\n{e}")
            finally:
                self.is_processing = False
        
        threading.Thread(target=generar, daemon=True).start()
    
    def _abrir_excel(self):
        """Abre el archivo Excel generado"""
        excel_path = os.path.abspath(self.excel_manager.excel_path)
        
        if os.path.exists(excel_path):
            try:
                os.startfile(excel_path)
                self._log("ℹ Excel abierto")
            except Exception as e:
                self._log(f"✗ Error al abrir Excel: {e}")
                messagebox.showerror("Error", f"No se pudo abrir el Excel:\n{e}")
        else:
            self._log("✗ El archivo Excel no existe")
            messagebox.showwarning(
                "Archivo no encontrado",
                "El archivo Excel no existe.\n\nGenera el Excel primero usando el botón 'Generar Excel'."
            )


def main():
    """Función principal para ejecutar la aplicación"""
    root = TkinterDnD.Tk()
    app = PagosExtractorGUI(root)
    
    # Centrar ventana
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Manejador de cierre
    def on_closing():
        if app.is_processing:
            if messagebox.askokcancel("Salir", "Hay un procesamiento en curso. ¿Deseas salir?"):
                root.destroy()
        else:
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    root.mainloop()


if __name__ == "__main__":
    main()
