#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Gestión de Pagos desde WhatsApp
Extrae registros de pagos de chats .txt y los gestiona en Excel
"""

import re
import os
import sys
import json
import logging
import time
from datetime import datetime
from typing import List, Dict, Tuple, Optional, Generator
import unicodedata

try:
    import pandas as pd
except ImportError:
    print("Error: pandas no está instalado. Ejecuta: pip install pandas")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl no está instalado. Ejecuta: pip install openpyxl")
    sys.exit(1)


class PaymentManager:
    """Gestiona el parsing, normalización y almacenamiento de pagos"""
    
    def __init__(self, excel_path="Pagos.xlsx"):
        self.excel_path = excel_path
        self.config_path = "config.json"
        self.setup_logging()
        self.load_config()
        
    def load_config(self):
        """Carga configuración desde config.json, si no existe el json, se crea uno por defecto"""
        self.config = {
            "horarios": {
                "matutino": "< 13:00",
                "vespertino": ">= 13:00",
                "archivo_procesado": None,
                "corte_actual": None
            },
            "mapeo_id_grupos": {}
        }
        
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            except Exception as e:
                logging.error(f"Error cargando config: {e}")
    
    def save_config(self):
        """Guarda configuración a config.json"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logging.error(f"Error guardando config: {e}")
    
    def get_group_info_from_config(self, payment_id: str) -> Tuple[Optional[str], Optional[str]]:
        """
        Obtiene nombre y sucursal normalizados desde config.json
        Retorna: (nombre_normalizado, sucursal)
        """
        if payment_id in self.config.get("mapeo_id_grupos", {}):
            grupo_info = self.config["mapeo_id_grupos"][payment_id]
            return grupo_info.get("nombre"), grupo_info.get("sucursal")
        return None, None
        
    def setup_logging(self):
        """Configura el logging a archivo"""
        logging.basicConfig(
            filename='log.txt',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
    def normalize_sucursal(self, text: str) -> str:
        """Quita acentos de las sucursales"""
        if not text or text.strip() == '':
            return "Sin especificar"
        text = text.strip()
        nfd = unicodedata.normalize('NFD', text)
        ascii_text = nfd.encode('ascii', 'ignore').decode('ascii')
        return ascii_text
    
    def normalize_number(self, text: str) -> float:
        """Normaliza números quitando $, comas y convirtiendo a float"""
        if not text:
            return 0.0
        cleaned = re.sub(r'[\$,\s]', '', str(text))
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    
    def get_current_corte(self) -> str:
        """
        Determina el corte horario actual basado en la hora del sistema
        Matutino < 13:00, Vespertino >= 13:00
        """
        hora_actual = datetime.now().hour
        corte = "Matutino" if hora_actual < 13 else "Vespertino"
        
        # Guardar en config
        self.config["horarios"]["corte_actual"] = corte
        self.save_config()
        
        return corte
    
    def extract_all_payments_from_lines(self, lines: List[str], filename: str, corte: str = None) -> List[Dict]:
        """Extrae todos los pagos de las líneas del archivo"""
        entries = []
        msg_pattern = r'\[(\d{2}/\d{2}/\d{2}), (\d{2}:\d{2}:\d{2})\] ([^:]+): (.+)'
        
        i = 0
        current_fecha = None
        current_hora = None
        
        while i < len(lines):
            line = lines[i]
            match = re.match(msg_pattern, line)
            
            if match:
                current_fecha = match.group(1)
                current_hora = match.group(2)
                content = match.group(4)
                
                # Acumular líneas siguientes hasta el siguiente mensaje
                following_lines = []
                j = i + 1
                while j < len(lines) and not re.match(msg_pattern, lines[j]):
                    following_lines.append(lines[j].strip())
                    j += 1
                
                # Combinar contenido
                full_content = content + '\n' + '\n'.join(following_lines)
                
                # Extraer grupos de este mensaje
                extracted = self.extract_payments_from_content(
                    full_content, current_fecha, current_hora, filename, corte
                )
                entries.extend(extracted)
                
                i = j
            else:
                i += 1
        
        return entries
    
    def extract_payments_from_content(self, content: str, fecha: str, hora: str, filename: str, corte: str = None) -> List[Dict]:
        """Extrae uno o más pagos del contenido de un mensaje"""
        entries = []
        
        # Ignorar mensajes del sistema
        if any(ignore in content for ignore in ['Creaste el grupo', 'cifrados de extremo a extremo']):
            return entries
        
        # Detectar tipo: Individual (Cliente) o Grupal (Grupo)
        es_individual = bool(re.search(r'\bCliente\b', content, re.IGNORECASE))
        es_grupal = bool(re.search(r'\bGrupo\b|\bGRUPO\b', content, re.IGNORECASE))
        
        # Si no hay ni Cliente ni Grupo, no procesar
        if not es_individual and not es_grupal:
            return entries
        
        # Si es individual, usar extract_single_payment
        if es_individual:
            single_entry = self.extract_single_payment(content, fecha, hora, filename, corte)
            if single_entry:
                entries.append(single_entry)
            return entries
        
        # Buscar todos los grupos en el contenido (solo para grupales)
        grupo_pattern = r'(?:Grupo|GRUPO)\s*:?\s*([A-Za-zÀ-ÿ\s]+?)(?:\s+0*\d{6})?\s+ID\s*:?\s*0*(\d{1,6})'
        grupo_matches = list(re.finditer(grupo_pattern, content, re.IGNORECASE))
        
        if not grupo_matches:
            # Intentar extraer un solo grupo
            single_entry = self.extract_single_payment(content, fecha, hora, filename, corte)
            if single_entry:
                entries.append(single_entry)
            return entries
        
        # Extraer datos para cada grupo encontrado
        for match in grupo_matches:
            try:
                grupo = match.group(1).strip()
                payment_id = match.group(2).zfill(6)
                
                # Extraer datos después del match de grupo
                start_pos = match.end()
                
                # Buscar Pago
                pago_match = re.search(r'Pago\s*:?\s*\$?\s*([\d,\.]+)', content[start_pos:])
                if not pago_match:
                    continue
                pago = self.normalize_number(pago_match.group(1))
                
                # Buscar Ahorro
                ahorro_match = re.search(r'Ahorro\s*:?\s*\$?\s*([\d,\.]+)', content[start_pos:])
                ahorro = self.normalize_number(ahorro_match.group(1)) if ahorro_match else 0.0
                
                # Buscar Sucursal
                sucursal_match = re.search(r'Sucursal\s*:?\s*([A-Za-zÀ-ÿ\s]+?)(?=\s*(?:N[úu]mero|$))', content[start_pos:])
                sucursal = sucursal_match.group(1).strip() if sucursal_match else None
                
                # Buscar Número de pago
                num_match = re.search(r'(?:Número de pago|N[úu]mero de pago|N pago|N Pago)\s*:?\s*(\d+)', content[start_pos:], re.IGNORECASE)
                num_pago = int(num_match.group(1)) if num_match else None
                
                # Intentar obtener info normalizada del config
                nombre_config, sucursal_config = self.get_group_info_from_config(payment_id)
                
                # Usar corte horario actual o determinar desde hora
                if corte is None:
                    hora_int = int(hora.split(':')[0]) if ':' in hora else 12
                    corte = "Matutino" if hora_int < 12 else "Vespertino"
                
                # Calcular Total
                total_calculado = round(pago + ahorro, 2)
                
                # Buscar Total en el contenido para validación
                total_match = re.search(r'Total\s*:?\s*\$?\s*([\d,\.]+)', content, re.IGNORECASE)
                if total_match:
                    total_dado = self.normalize_number(total_match.group(1))
                    # Validar que Total = Pago + Ahorro (tolerancia 0.01)
                    if abs(total_dado - total_calculado) > 0.01:
                        logging.warning(f"Discrepancia en Total para ID {payment_id}: "
                                      f"Calculado={total_calculado}, Dado={total_dado}, "
                                      f"Diferencia={abs(total_dado - total_calculado)}")
                
                entry = {
                    'Tipo': 'Gpo',  # Es grupal
                    'ID': payment_id,
                    'Grupo': nombre_config if nombre_config else grupo.upper(),
                    'Fecha': fecha,
                    'Hora': hora,
                    'Pago': round(pago, 2),
                    'Ahorro': round(ahorro, 2),
                    'Total': total_calculado,
                    'Número de Pago': num_pago,
                    'Sucursal': sucursal_config if sucursal_config else (self.normalize_sucursal(sucursal) if sucursal else "Sin especificar"),
                    'Corte': corte,
                    'Confirmado': 'No',
                    'Archivo': filename
                }
                
                entries.append(entry)
            except Exception as e:
                logging.error(f"Error parseando entrada: {e}")
                continue
        
        return entries
    
    def extract_single_payment(self, content: str, fecha: str, hora: str, filename: str, corte: str = None) -> Optional[Dict]:
        """Extrae un solo pago del contenido (Individual o Grupal)"""
        # Detectar tipo: Individual (Cliente) o Grupal (Grupo)
        es_individual = bool(re.search(r'\bCliente\b', content, re.IGNORECASE))
        es_grupal = bool(re.search(r'\bGrupo\b|\bGRUPO\b', content, re.IGNORECASE))
        
        if not es_individual and not es_grupal:
            return None
        
        # Buscar ID (requerido en ambos casos)
        id_match = re.search(r'ID\s*:?\s*0*(\d{1,6})', content)
        if not id_match:
            return None
        payment_id = id_match.group(1).zfill(6)
        
        # Buscar Pago (requerido en ambos casos)
        pago_match = re.search(r'Pago\s*:?\s*\$?\s*([\d,\.]+)', content)
        if not pago_match:
            return None
        pago = self.normalize_number(pago_match.group(1))
        
        # Buscar Sucursal
        sucursal_match = re.search(r'Sucursal\s*:?\s*([A-Za-zÀ-ÿ\s]+?)(?=\s*(?:N[úu]mero|$))', content)
        sucursal = sucursal_match.group(1).strip() if sucursal_match else None
        
        # Intentar obtener info normalizada del config
        nombre_config, sucursal_config = self.get_group_info_from_config(payment_id)
        
        # Usar corte horario actual o determinar desde hora
        if corte is None:
            hora_int = int(hora.split(':')[0]) if ':' in hora else 12
            corte = "Matutino" if hora_int < 12 else "Vespertino"
        
        # Procesar según tipo
        if es_individual:
            # INDIVIDUAL: Buscar nombre del cliente
            cliente_match = re.search(r'Cliente\s+([A-Za-zÀ-ÿ\s]+?)(?=\s+ID)', content, re.IGNORECASE)
            if not cliente_match:
                return None
            cliente_nombre = cliente_match.group(1).strip()
            
            # Reglas para Individual
            grupo = nombre_config if nombre_config else cliente_nombre.upper()
            ahorro = 0.0
            total_calculado = round(pago, 2)  # Total = Pago
            num_pago = None  # Sin número de pago
            tipo = 'Ind'
        else:
            # GRUPAL: Buscar Grupo
            grupo_match = re.search(r'(?:Grupo|GRUPO)\s*:?\s*([A-Za-zÀ-ÿ\s]+?)(?=\s*ID|\s*\d{6})', content)
            if not grupo_match:
                return None
            grupo = grupo_match.group(1).strip()
            
            # Buscar Ahorro (solo para grupales)
            ahorro_match = re.search(r'Ahorro\s*:?\s*\$?\s*([\d,\.]+)', content)
            ahorro = self.normalize_number(ahorro_match.group(1)) if ahorro_match else 0.0
            
            # Buscar Número de pago (solo para grupales)
            num_match = re.search(r'(?:Número de pago|N[úu]mero de pago|N pago|N Pago)\s*:?\s*(\d+)', content, re.IGNORECASE)
            num_pago = int(num_match.group(1)) if num_match else None
            
            # Calcular Total para grupal
            total_calculado = round(pago + ahorro, 2)
            tipo = 'Gpo'
        
        # Buscar Total en el contenido para validación (solo grupal puede tenerlo)
        if es_grupal:
            total_match = re.search(r'Total\s*:?\s*\$?\s*([\d,\.]+)', content, re.IGNORECASE)
            if total_match:
                total_dado = self.normalize_number(total_match.group(1))
                # Validar que Total = Pago + Ahorro (tolerancia 0.01)
                if abs(total_dado - total_calculado) > 0.01:
                    logging.warning(f"Discrepancia en Total para ID {payment_id}: "
                                  f"Calculado={total_calculado}, Dado={total_dado}, "
                                  f"Diferencia={abs(total_dado - total_calculado)}")
        
        return {
            'Tipo': tipo,
            'ID': payment_id,
            'Grupo': nombre_config if nombre_config else grupo.upper(),
            'Fecha': fecha,
            'Hora': hora,
            'Pago': round(pago, 2),
            'Ahorro': round(ahorro, 2),
            'Total': total_calculado,
            'Número de Pago': num_pago,
            'Sucursal': sucursal_config if sucursal_config else (self.normalize_sucursal(sucursal) if sucursal else "Sin especificar"),
            'Corte': corte,
            'Confirmado': 'No',
            'Archivo': filename
        }
    
    def get_last_timestamp(self) -> Optional[str]:
        """Obtiene el último timestamp procesado desde la hoja Meta"""
        try:
            if not os.path.exists(self.excel_path):
                return None
            
            df_meta = pd.read_excel(self.excel_path, sheet_name='Meta', engine='openpyxl')
            if df_meta.empty or 'ultimo_timestamp' not in df_meta.columns:
                return None
            
            return df_meta.iloc[0]['ultimo_timestamp']
        except Exception as e:
            logging.error(f"Error leyendo último timestamp: {e}")
            return None
    
    def save_timestamp(self, timestamp: str):
        """Guarda el último timestamp procesado en la hoja Meta"""
        try:
            wb = openpyxl.load_workbook(self.excel_path) if os.path.exists(self.excel_path) else openpyxl.Workbook()
            
            # Eliminar hoja Meta si existe
            if 'Meta' in wb.sheetnames:
                wb.remove(wb['Meta'])
            
            ws_meta = wb.create_sheet('Meta', 0)
            ws_meta.append(['ultimo_timestamp'])
            ws_meta.append([timestamp])
            ws_meta.sheet_state = 'hidden'
            
            wb.save(self.excel_path)
            wb.close()
        except Exception as e:
            logging.error(f"Error guardando timestamp: {e}")
    
    def extract_last_timestamp_from_file(self, filepath: str) -> Optional[str]:
        """Extrae el timestamp del último mensaje en el archivo"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            msg_pattern = r'\[(\d{2}/\d{2}/\d{2}), (\d{2}:\d{2}:\d{2})\]'
            for line in reversed(lines):
                match = re.search(msg_pattern, line)
                if match:
                    fecha = match.group(1)
                    hora = match.group(2)
                    dd, mm, yy = fecha.split('/')
                    timestamp = f"{yy}/{mm}/{dd} {hora}"
                    return timestamp
            return None
        except Exception as e:
            logging.error(f"Error extrayendo timestamp de {filepath}: {e}")
            return None
    
    def process_file(self, filepath: str) -> Tuple[List[Dict], int, int]:
        """Procesa un archivo .txt y extrae pagos"""
        entries = []
        errors = 0
        duplicates = 0
        
        # Verificar si el archivo ya fue procesado
        last_ts = self.extract_last_timestamp_from_file(filepath)
        if last_ts:
            stored_ts = self.get_last_timestamp()
            if stored_ts and last_ts <= stored_ts:
                logging.info(f"Archivo {filepath} ya procesado")
                return [], 0, 1
        
        # Obtener corte horario actual
        corte_actual = self.get_current_corte()
        
        try:
            # Usar generador línea por línea en lugar de readlines
            lines = []
            with open(filepath, 'r', encoding='utf-8') as f:
                for line in f:
                    lines.append(line.rstrip('\n\r'))
            
            filename = os.path.basename(filepath)
            entries = self.extract_all_payments_from_lines(lines, filename, corte_actual)
            
            # Eliminar duplicados usando ID + Grupo + Pago + Ahorro + timestamp
            seen = set()
            unique_entries = []
            for entry in entries:
                key = f"{entry['ID']}_{entry['Grupo']}_{entry['Pago']}_{entry['Ahorro']}_{entry['Fecha']} {entry['Hora']}"
                if key not in seen:
                    seen.add(key)
                    unique_entries.append(entry)
                else:
                    duplicates += 1
            
            entries = unique_entries
            
            # Guardar timestamp si se procesó exitosamente
            if entries and last_ts:
                self.save_timestamp(last_ts)
                logging.info(f"Procesados {len(entries)} pagos de {filepath}")
                
                # Guardar corte de procesamiento en config
                self.config["horarios"]["archivo_procesado"] = corte_actual
                self.save_config()
            
        except Exception as e:
            logging.error(f"Error procesando {filepath}: {e}")
            errors += 1
        
        return entries, errors, duplicates
    
    def add_to_excel(self, entries: List[Dict]) -> int:
        """Agrega entradas al Excel"""
        if not entries:
            logging.info("No hay entradas para agregar")
            return 0
        
        try:
            logging.info(f"Creando DataFrame con {len(entries)} entradas")
            df_new = pd.DataFrame(entries)
            
            # Orden EXACTO de columnas con 'Tipo' como primera columna
            cols_orden = ['Tipo', 'ID', 'Grupo', 'Fecha', 'Hora', 'Pago', 'Ahorro', 'Total', 
                         'Número de Pago', 'Sucursal', 'Corte', 'Confirmado']
            
            # Eliminar 'Archivo' que no debe ir al Excel
            if 'Archivo' in df_new.columns:
                df_new = df_new.drop(columns=['Archivo'])
            
            # Asegurar que todas las columnas existan (rellenar con valores por defecto si faltan)
            for col in cols_orden:
                if col not in df_new.columns:
                    if col == 'Tipo':
                        # Si no hay Tipo, inferir de otros campos
                        df_new[col] = df_new.apply(
                            lambda row: 'Gpo' if pd.notna(row.get('Ahorro', 0)) and float(row.get('Ahorro', 0)) > 0 
                                       else 'Ind', axis=1
                        )
                    else:
                        df_new[col] = None
            
            # Reordenar columnas al orden exacto especificado
            df_new = df_new.reindex(columns=cols_orden)
            
            # Convertir ID a string y asegurar formato de 6 dígitos con ceros a la izquierda
            if 'ID' in df_new.columns:
                df_new['ID'] = df_new['ID'].astype(str).str.zfill(6)
            
            # Verificar si existe archivo con hoja Pagos
            df_existing = None
            if os.path.exists(self.excel_path):
                try:
                    df_existing = pd.read_excel(self.excel_path, sheet_name='Pagos', engine='openpyxl')
                    
                    # Convertir ID existente a string y rellenar con ceros a la izquierda
                    if 'ID' in df_existing.columns:
                        df_existing['ID'] = df_existing['ID'].astype(str).str.replace('.0', '', regex=False)
                        df_existing['ID'] = df_existing['ID'].str.zfill(6)
                    
                    # Si Excel existente no tiene 'Tipo', agregarlo y rellenar
                    if 'Tipo' not in df_existing.columns:
                        # Inferir Tipo de campos existentes
                        df_existing['Tipo'] = df_existing.apply(
                            lambda row: 'Gpo' if pd.notna(row.get('Ahorro', 0)) and float(row.get('Ahorro', 0)) > 0 
                                       else 'Ind', axis=1
                        )
                    
                    # Asegurar todas las columnas del orden especificado
                    for col in cols_orden:
                        if col not in df_existing.columns:
                            df_existing[col] = None
                    
                    # Reordenar columnas existentes al orden exacto
                    df_existing = df_existing.reindex(columns=cols_orden)
                except Exception as e:
                    logging.warning(f"Error leyendo Excel existente: {e}")
                    df_existing = None
            
            if df_existing is not None and not df_existing.empty:
                df_final = pd.concat([df_existing, df_new]).drop_duplicates(
                    subset=['ID', 'Grupo', 'Pago', 'Ahorro'], keep='first'
                ).reset_index(drop=True)
            else:
                df_final = df_new
            
            logging.info(f"Guardando {len(df_final)} registros en {self.excel_path}")
            
            # Crear ExcelWriter y agregar ambas hojas
            with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
                df_final.to_excel(writer, sheet_name='Pagos', index=False)
                # Crear hoja Meta vacía
                df_meta = pd.DataFrame({'ultimo_timestamp': ['']})
                df_meta.to_excel(writer, sheet_name='Meta', index=False)
            
            # Configurar formato de Excel y ocultar Meta con retries
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    wb = openpyxl.load_workbook(self.excel_path)
                    
                    # Configurar columna ID como texto para preservar formato
                    if 'Pagos' in wb.sheetnames:
                        ws = wb['Pagos']
                        for cell in ws[1]:  # Primera fila (encabezados)
                            if cell.value == 'ID':
                                col_letter = cell.column_letter
                                # Formatear todas las celdas de la columna ID como texto
                                for row in range(2, ws.max_row + 1):
                                    ws[f'{col_letter}{row}'].number_format = '@'  # @ = texto
                    
                    # Ocultar hoja Meta
                    if 'Meta' in wb.sheetnames and len(wb.sheetnames) > 1:
                        meta_idx = wb.sheetnames.index('Meta')
                        wb.worksheets[meta_idx].sheet_state = 'hidden'
                    
                    wb.save(self.excel_path)
                    wb.close()
                    break
                except PermissionError as pe:
                    if attempt < max_retries - 1:
                        logging.warning(f"Intento {attempt + 1} de {max_retries}: Permiso denegado. Esperando...")
                        time.sleep(1)
                    else:
                        logging.error(f"NO se pudo guardar Excel tras {max_retries} intentos. Cierra el archivo en Excel.")
                        raise
                except Exception as meta_error:
                    logging.warning(f"No se pudo configurar formato del Excel: {meta_error}")
                    break
            
            logging.info(f"Guardado exitoso: {len(df_final)} registros")
            return len(df_final)
        except Exception as e:
            import traceback
            logging.error(f"Error agregando a Excel: {e}")
            logging.error(traceback.format_exc())
            return 0
    
    def process_confirmations(self, filepath: str) -> Tuple[List[Dict], List[str]]:
        """
        Procesa archivo de confirmaciones y actualiza registros en Excel
        Retorna: (lista de confirmaciones procesadas, lista de alertas de no encontrados)
        """
        alerts = []
        confirmed_entries = []
        
        # Procesar archivo de confirmaciones directamente (sin filtro de timestamp)
        # Obtener corte horario actual para las confirmaciones
        corte_actual = self.get_current_corte()
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = []
                for line in f:
                    lines.append(line.rstrip('\n\r'))
            
            filename = os.path.basename(filepath)
            entries = self.extract_all_payments_from_lines(lines, filename, corte_actual)
            
            # Eliminar duplicados
            seen = set()
            unique_entries = []
            for entry in entries:
                key = f"{entry['ID']}_{entry['Grupo']}_{entry['Pago']}_{entry['Ahorro']}"
                if key not in seen:
                    seen.add(key)
                    unique_entries.append(entry)
            
            entries = unique_entries
        except Exception as e:
            logging.error(f"Error leyendo confirmaciones: {e}")
            alerts.append(f"Error leyendo archivo de confirmaciones: {e}")
        
        if not entries:
            alerts.append("No se encontraron confirmaciones válidas en el archivo")
            return [], alerts
        
        # Verificar que existe el Excel
        if not os.path.exists(self.excel_path):
            alerts.append("No existe archivo de pagos para confirmar")
            return [], alerts
        
        try:
            # Leer hoja de Pagos
            df_pagos = pd.read_excel(self.excel_path, sheet_name='Pagos', engine='openpyxl')
            
            for conf_entry in entries:
                match_found = False
                conf_tipo = conf_entry.get('Tipo', 'Gpo')  # Por defecto Gpo si no viene
                logging.info(f"Buscando confirmación: Tipo={conf_tipo}, ID={conf_entry['ID']}, Grupo={conf_entry['Grupo']}, "
                           f"Pago={conf_entry['Pago']}, Ahorro={conf_entry['Ahorro']}")
                
                # Buscar coincidencia en df_pagos con Tipo + ID + Grupo + Pago + Ahorro
                for idx in df_pagos.index:
                    row = df_pagos.iloc[idx]
                    
                    # Comparar Tipo
                    excel_tipo = str(row.get('Tipo', 'Gpo')).strip() if pd.notna(row.get('Tipo')) else 'Gpo'
                    if excel_tipo != conf_tipo:
                        continue
                    
                    # Convertir ID a string y rellenar con ceros para comparación
                    excel_id = str(row['ID']).replace('.0', '').zfill(6) if pd.notna(row['ID']) else ''
                    conf_id = str(conf_entry['ID']).strip().zfill(6)
                    
                    # Comparar ID
                    if excel_id != conf_id:
                        continue
                    
                    # Comparar Grupo (case-insensitive)
                    excel_grupo = str(row['Grupo']).strip().upper() if pd.notna(row['Grupo']) else ''
                    conf_grupo = str(conf_entry['Grupo']).strip().upper()
                    if excel_grupo != conf_grupo:
                        continue
                    
                    # Comparar Pago con tolerancia 0.01
                    excel_pago = float(row['Pago']) if pd.notna(row['Pago']) else 0.0
                    conf_pago = float(conf_entry['Pago'])
                    if abs(excel_pago - conf_pago) > 0.01:
                        logging.warning(f"Discrepancia en Pago para ID {conf_id}: "
                                      f"Excel={excel_pago} vs Confirmación={conf_pago}")
                        continue
                    
                    # Comparar Ahorro con tolerancia 0.01
                    excel_ahorro = float(row['Ahorro']) if pd.notna(row['Ahorro']) else 0.0
                    conf_ahorro = float(conf_entry['Ahorro'])
                    if abs(excel_ahorro - conf_ahorro) > 0.01:
                        logging.warning(f"Discrepancia en Ahorro para ID {conf_id}: "
                                      f"Excel={excel_ahorro} vs Confirmación={conf_ahorro}")
                    
                    # Match completo encontrado
                    logging.info(f"MATCH ENCONTRADO: Tipo={excel_tipo}, ID={conf_id}, Grupo={conf_grupo}")
                    match_found = True
                    
                    # Actualizar a "Sí" en columna Confirmado
                    df_pagos.at[idx, 'Confirmado'] = 'Sí'
                    
                    # Actualizar Ahorro si difiere
                    if abs(excel_ahorro - conf_ahorro) > 0.01:
                        df_pagos.at[idx, 'Ahorro'] = conf_ahorro
                        df_pagos.at[idx, 'Total'] = excel_pago + conf_ahorro
                    
                    # Copiar registro completo para hoja de confirmados
                    confirmed_entry = df_pagos.iloc[idx].to_dict()
                    confirmed_entries.append(confirmed_entry)
                    
                    break
                
                if not match_found:
                    alerts.append(
                        f"No se encontró: ID {conf_entry['ID']}, Grupo {conf_entry['Grupo']}, "
                        f"Pago {conf_entry['Pago']}, Ahorro {conf_entry['Ahorro']}"
                    )
            
            # Guardar cambios en hoja Pagos
            with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_pagos.to_excel(writer, sheet_name='Pagos', index=False)
            
            # Actualizar hoja Pagos Confirmados
            if confirmed_entries:
                df_confirmed = pd.DataFrame(confirmed_entries)
                
                # Intentar leer confirmados existentes
                try:
                    df_existing_confirmed = pd.read_excel(
                        self.excel_path, sheet_name='Pagos Confirmados', engine='openpyxl'
                    )
                    # Combinar con los nuevos
                    df_confirmed = pd.concat([df_existing_confirmed, df_confirmed])
                except:
                    pass
                
                # Guardar hoja de confirmados
                with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_confirmed.to_excel(writer, sheet_name='Pagos Confirmados', index=False)
                
                logging.info(f"Confirmados {len(confirmed_entries)} pagos")
            
        except Exception as e:
            import traceback
            logging.error(f"Error procesando confirmaciones: {e}")
            logging.error(traceback.format_exc())
            alerts.append(f"Error procesando confirmaciones: {str(e)}")
        
        return confirmed_entries, alerts
    
    def clear_all_data(self) -> bool:
        """
        Limpia todos los registros del sistema
        Elimina Excel, limpia config y log
        Retorna True si se limpió exitosamente
        """
        errors = []
        
        # Eliminar archivo Excel con retries
        if os.path.exists(self.excel_path):
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    os.remove(self.excel_path)
                    logging.info(f"Eliminado archivo {self.excel_path}")
                    break
                except PermissionError as pe:
                    if attempt < max_retries - 1:
                        logging.warning(f"Intento {attempt + 1} de {max_retries}: Permiso denegado. Esperando...")
                        time.sleep(1)
                    else:
                        errors.append(f"NO se pudo eliminar {self.excel_path} tras {max_retries} intentos (archivo abierto en Excel)")
                except Exception as e:
                    errors.append(f"Error eliminando Excel: {e}")
                    break
        
        # Limpiar config.json
        try:
            self.config = {
                "horarios": {
                    "matutino": "< 13:00",
                    "vespertino": ">= 13:00",
                    "archivo_procesado": None,
                    "corte_actual": None
                },
                "mapeo_id_grupos": {}
            }
            self.save_config()
            logging.info("Config.json limpiado")
        except Exception as e:
            errors.append(f"Error limpiando config: {e}")
        
        # Limpiar log.txt
        try:
            if os.path.exists('log.txt'):
                os.remove('log.txt')
                logging.info("Log.txt eliminado")
        except Exception as e:
            errors.append(f"Error eliminando log: {e}")
        
        # Configurar logging de nuevo
        self.setup_logging()
        
        if errors:
            for error in errors:
                logging.warning(error)
            return False
        else:
            logging.info("Todos los datos fueron limpiados exitosamente")
            return True


def main():
    """Función principal para probar el script"""
    print("Sistema de Gestión de Pagos desde WhatsApp")
    print("=" * 50)
    
    manager = PaymentManager()
    
    # ============================================
    # OPCION PARA LIMPIAR REGISTROS (DESCOMENTAR SI SE NECESITA)
    # ============================================
    print("\nLimpiando todos los registros...")
    if manager.clear_all_data():
        print("[OK] Registros limpiados exitosamente")
    else:
        print("[ADVERTENCIA] Algunos archivos no pudieron ser eliminados")
        print("(Por ejemplo, Pagos.xlsx podría estar abierto en Excel)")
        print("El sistema continuará igualmente...")
    # ============================================
    
    # Procesar archivo de ejemplo
    filepath = "ejemplos/_chat.txt"
    
    if not os.path.exists(filepath):
        print(f"Error: No se encuentra el archivo {filepath}")
        return
    
    print(f"Procesando {filepath}...")
    entries, errors, duplicates = manager.process_file(filepath)
    
    print(f"\nResultados:")
    print(f"  Entradas extraídas: {len(entries)}")
    print(f"  Errores: {errors}")
    print(f"  Duplicados: {duplicates}")
    
    if entries:
        print(f"\nPrimeros 5 registros:")
        for i, entry in enumerate(entries[:5], 1):
            print(f"\n{i}. ID: {entry['ID']}")
            print(f"   Grupo: {entry['Grupo']}")
            print(f"   Fecha: {entry['Fecha']} {entry['Hora']}")
            print(f"   Pago: ${entry['Pago']}, Ahorro: ${entry['Ahorro']}, Total: ${entry['Total']}")
            print(f"   Sucursal: {entry['Sucursal']}")
            if entry['Número de Pago']:
                print(f"   Número de pago: {entry['Número de Pago']}")
        
        # Guardar en Excel
        print(f"\nGuardando en Excel...")
        added = manager.add_to_excel(entries)
        print(f"Guardados {added} registros en {manager.excel_path}")
        print(f"Total de entradas extraídas: {len(entries)}")
        print(f"Puedes abrir el archivo Excel para ver los resultados")
    else:
        print("No se encontraron pagos válidos")


if __name__ == "__main__":
    main()

