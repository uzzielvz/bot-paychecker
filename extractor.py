import re
import json
import hashlib
from datetime import datetime
from typing import Optional, Dict, List


class WhatsAppExtractor:
    """Extrae y procesa mensajes de pagos desde archivos de chat de WhatsApp"""
    
    def __init__(self, config_path: str = "config.json"):
        """Inicializa el extractor cargando la configuración"""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.cortes_horarios = self.config['cortes_horarios']
        
        # Patrón para detectar líneas de mensaje de WhatsApp
        # Formato real: [D/M/YY, H:MM:SS] Remitente: Mensaje
        # También soporta formato sin corchetes para compatibilidad
        self.whatsapp_pattern = re.compile(
            r'^\[?(\d{1,2}/\d{1,2}/\d{2,4}),\s+(\d{1,2}:\d{2}(?::\d{2})?(?:\s+[ap]\.\s+m\.)?)\]?\s+([^:]+):\s+(.*)$'
        )
        
        # Patrón 1: Formato compacto (una línea)
        # Ejemplo: Nueva Luz 000031/ Puebla/ pago $5,852/ ahorro/ $ 348
        # También: Magia 000055/ puebla/ $9,045/ $443 (sin palabras pago/ahorro)
        self.pattern_compact = re.compile(
            r'([A-Za-záéíóúñÑ\s]+?)\s*(\d{6})\s*/\s*([A-Za-záéíóúñÑ]+)\s*/\s*(?:[Pp]ago\s*)?\$?\s*([\d,]+)\s*/?\s*(?:[Aa]horro\s*)?\$?\s*([\d,]+)',
            re.IGNORECASE
        )
        
        # Patrón adicional para formato: "Nombre ID pago ahorro" (sin separadores /)
        # Ejemplo: Bienvenidos 000094 12 920.11 pago 775.89 ahorro
        self.pattern_compact_simple = re.compile(
            r'([A-Za-záéíóúñÑ\s]+?)\s*(\d{6})\s+(?:\d+\s+)?([\d,]+\.?\d*)\s+[Pp]ago\s+([\d,]+\.?\d*)\s+[Aa]horro',
            re.IGNORECASE
        )
        
        # Patrón 2: Formato multilínea (buffer de 4 líneas)
        # Ejemplo: Grupo CATALEYA \n ID000080 \n PAGO 7606 \n AHORRO 644
        self.pattern_multiline_grupo = re.compile(r'[Gg]rupo\s+([A-Za-záéíóúñÑ\s]+)', re.IGNORECASE)
        self.pattern_multiline_id = re.compile(r'ID\s*(\d{6})', re.IGNORECASE)
        self.pattern_multiline_pago = re.compile(r'PAGO\s*\$?\s*([\d,]+)', re.IGNORECASE)
        self.pattern_multiline_ahorro = re.compile(r'AHORRO\s*\$?\s*([\d,]+)', re.IGNORECASE)
        
    def es_mensaje_sistema(self, contenido: str) -> bool:
        """Identifica y filtra mensajes del sistema de WhatsApp"""
        mensajes_sistema = [
            'cifrados de extremo a extremo',
            'mensajes temporales',
            'archivo adjunto',
            'Se eliminó este mensaje',
            'Multimedia omitido',
            '<Multimedia omitido>',
            '.opus',
            '.jpg',
            '.png',
            '.pdf',
            '.zip',
            '.csv',
            '.xlsx'
        ]
        
        for msg_sistema in mensajes_sistema:
            if msg_sistema.lower() in contenido.lower():
                return True
        
        return False
    
    def parsear_linea_whatsapp(self, linea: str) -> Optional[Dict]:
        """Parsea una línea de chat de WhatsApp y extrae fecha, hora, remitente y mensaje"""
        match = self.whatsapp_pattern.match(linea)
        if not match:
            return None
        
        fecha_str, hora_str, remitente, contenido = match.groups()
        
        # Convertir fecha y hora a datetime
        try:
            # Detectar si tiene año corto (YY) o largo (YYYY)
            partes_fecha = fecha_str.split('/')
            if len(partes_fecha[2]) == 2:
                # Año corto, convertir a 4 dígitos
                año_corto = int(partes_fecha[2])
                año_completo = 2000 + año_corto if año_corto < 50 else 1900 + año_corto
                fecha_str = f"{partes_fecha[0]}/{partes_fecha[1]}/{año_completo}"
            
            # Detectar formato de hora
            if 'a. m.' in hora_str or 'p. m.' in hora_str:
                # Formato 12 horas con a.m./p.m.
                fecha_hora_str = f"{fecha_str} {hora_str}"
                fecha_hora_str = fecha_hora_str.replace('a. m.', 'AM').replace('p. m.', 'PM')
                fecha_hora = datetime.strptime(fecha_hora_str, '%d/%m/%Y %I:%M %p')
            else:
                # Formato 24 horas (puede incluir segundos o no)
                if hora_str.count(':') == 2:
                    # Con segundos: 17:13:49
                    fecha_hora = datetime.strptime(f"{fecha_str} {hora_str}", '%d/%m/%Y %H:%M:%S')
                else:
                    # Sin segundos: 17:13
                    fecha_hora = datetime.strptime(f"{fecha_str} {hora_str}", '%d/%m/%Y %H:%M')
        except (ValueError, IndexError) as e:
            return None
        
        return {
            'fecha_hora': fecha_hora,
            'remitente': remitente.strip(),
            'contenido': contenido.strip()
        }
    
    def determinar_corte(self, hora: datetime) -> str:
        """Determina el corte horario según la hora del mensaje"""
        hora_time = hora.time()
        
        for corte in self.cortes_horarios:
            inicio = datetime.strptime(corte['hora_inicio'], '%H:%M').time()
            fin = datetime.strptime(corte['hora_fin'], '%H:%M').time()
            
            if inicio <= hora_time <= fin:
                return corte['nombre']
        
        return "Sin Corte"
    
    def limpiar_numero(self, numero_str: str) -> float:
        """Limpia y convierte string numérico a float (elimina $, comas, espacios, asteriscos)"""
        numero_limpio = numero_str.replace('$', '').replace(',', '').replace(' ', '').replace('*', '').strip()
        try:
            return float(numero_limpio)
        except ValueError:
            return 0.0
    
    def limpiar_texto(self, texto: str) -> str:
        """Limpia texto eliminando asteriscos de markdown y espacios extra"""
        return texto.replace('*', '').strip()
    
    def extraer_pago_compacto(self, contenido: str) -> Optional[Dict]:
        """Extrae datos de pago en formato compacto (una línea)"""
        # Intentar formato con separadores /
        match = self.pattern_compact.search(contenido)
        if match:
            grupo, id_grupo, sucursal, pago_str, ahorro_str = match.groups()
            return {
                'grupo': self.limpiar_texto(grupo),
                'id_grupo': id_grupo.strip(),
                'sucursal': self.limpiar_texto(sucursal),
                'pago': self.limpiar_numero(pago_str),
                'ahorro': self.limpiar_numero(ahorro_str)
            }
        
        # Intentar formato simple: "Nombre ID pago ahorro"
        match = self.pattern_compact_simple.search(contenido)
        if match:
            grupo, id_grupo, pago_str, ahorro_str = match.groups()
            return {
                'grupo': self.limpiar_texto(grupo),
                'id_grupo': id_grupo.strip(),
                'sucursal': 'N/A',
                'pago': self.limpiar_numero(pago_str),
                'ahorro': self.limpiar_numero(ahorro_str)
            }
        
        return None
    
    def extraer_pago_multilinea(self, lineas_buffer: List[str]) -> Optional[Dict]:
        """Extrae datos de pago en formato multilínea (múltiples líneas)"""
        # Unir todas las líneas del buffer en un solo string
        contenido_completo = '\n'.join(lineas_buffer)
        
        # Buscar cada componente
        match_grupo = self.pattern_multiline_grupo.search(contenido_completo)
        match_id = self.pattern_multiline_id.search(contenido_completo)
        match_pago = self.pattern_multiline_pago.search(contenido_completo)
        match_ahorro = self.pattern_multiline_ahorro.search(contenido_completo)
        
        # Si encontramos al menos grupo, id y pago, es válido
        if match_grupo and match_id and match_pago:
            grupo = self.limpiar_texto(match_grupo.group(1))
            id_grupo = match_id.group(1).strip()
            pago = self.limpiar_numero(match_pago.group(1))
            ahorro = self.limpiar_numero(match_ahorro.group(1)) if match_ahorro else 0.0
            
            return {
                'grupo': grupo,
                'id_grupo': id_grupo,
                'sucursal': 'N/A',  # No siempre está en formato multilínea
                'pago': pago,
                'ahorro': ahorro
            }
        
        return None
    
    def generar_hash(self, fecha_hora: datetime, remitente: str, contenido: str) -> str:
        """Genera un hash único para identificar mensajes duplicados"""
        cadena = f"{fecha_hora.isoformat()}|{remitente}|{contenido}"
        return hashlib.sha256(cadena.encode('utf-8')).hexdigest()
    
    def procesar_archivo(self, ruta_archivo: str) -> List[Dict]:
        """Procesa un archivo de chat de WhatsApp y extrae todos los pagos"""
        pagos_extraidos = []
        
        with open(ruta_archivo, 'r', encoding='utf-8') as f:
            lineas = f.readlines()
        
        # Buffer para procesar mensajes multilínea
        buffer_multilinea = []
        mensaje_actual = None
        
        for i, linea in enumerate(lineas):
            linea = linea.strip()
            if not linea:
                continue
            
            # Intentar parsear como línea de WhatsApp
            parsed = self.parsear_linea_whatsapp(linea)
            
            if parsed:
                # Es una nueva línea de mensaje
                # Procesar buffer anterior si existe
                if buffer_multilinea and mensaje_actual:
                    pago_multi = self.extraer_pago_multilinea(buffer_multilinea)
                    if pago_multi:
                        pago_multi['fecha_hora'] = mensaje_actual['fecha_hora']
                        pago_multi['remitente'] = mensaje_actual['remitente']
                        pago_multi['mensaje_original'] = '\n'.join(buffer_multilinea)
                        pago_multi['corte_horario'] = self.determinar_corte(mensaje_actual['fecha_hora'])
                        pago_multi['hash'] = self.generar_hash(
                            mensaje_actual['fecha_hora'],
                            mensaje_actual['remitente'],
                            pago_multi['mensaje_original']
                        )
                        pagos_extraidos.append(pago_multi)
                
                # Limpiar buffer
                buffer_multilinea = []
                mensaje_actual = parsed
                
                # Filtrar mensajes del sistema
                if self.es_mensaje_sistema(parsed['contenido']):
                    mensaje_actual = None
                    continue
                
                # Intentar extraer pago en formato compacto
                pago_compacto = self.extraer_pago_compacto(parsed['contenido'])
                if pago_compacto:
                    pago_compacto['fecha_hora'] = parsed['fecha_hora']
                    pago_compacto['remitente'] = parsed['remitente']
                    pago_compacto['mensaje_original'] = parsed['contenido']
                    pago_compacto['corte_horario'] = self.determinar_corte(parsed['fecha_hora'])
                    pago_compacto['hash'] = self.generar_hash(
                        parsed['fecha_hora'],
                        parsed['remitente'],
                        parsed['contenido']
                    )
                    pagos_extraidos.append(pago_compacto)
                    mensaje_actual = None  # Ya procesado
                else:
                    # Agregar al buffer para procesamiento multilínea
                    buffer_multilinea.append(parsed['contenido'])
            else:
                # Continuación de mensaje anterior (multilínea)
                if mensaje_actual:
                    buffer_multilinea.append(linea)
        
        # Procesar último buffer si existe
        if buffer_multilinea and mensaje_actual:
            pago_multi = self.extraer_pago_multilinea(buffer_multilinea)
            if pago_multi:
                pago_multi['fecha_hora'] = mensaje_actual['fecha_hora']
                pago_multi['remitente'] = mensaje_actual['remitente']
                pago_multi['mensaje_original'] = '\n'.join(buffer_multilinea)
                pago_multi['corte_horario'] = self.determinar_corte(mensaje_actual['fecha_hora'])
                pago_multi['hash'] = self.generar_hash(
                    mensaje_actual['fecha_hora'],
                    mensaje_actual['remitente'],
                    pago_multi['mensaje_original']
                )
                pagos_extraidos.append(pago_multi)
        
        return pagos_extraidos


if __name__ == "__main__":
    # Prueba básica
    extractor = WhatsAppExtractor()
    print("Extractor de WhatsApp inicializado correctamente")
    print(f"Cortes horarios configurados: {len(extractor.cortes_horarios)}")

