import sqlite3
import json
from datetime import datetime
from typing import List, Dict, Optional
import os


class DatabaseManager:
    """Gestiona la base de datos SQLite para almacenar pagos"""
    
    def __init__(self, config_path: str = "config.json"):
        """Inicializa la conexión a la base de datos"""
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        db_dir = config['rutas']['database']
        os.makedirs(db_dir, exist_ok=True)
        
        self.db_path = os.path.join(db_dir, 'pagos.db')
        self.inicializar_base_datos()
    
    def inicializar_base_datos(self):
        """Crea las tablas necesarias si no existen"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Tabla principal de pagos
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pagos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha_mensaje TIMESTAMP NOT NULL,
                fecha_procesamiento TIMESTAMP NOT NULL,
                corte_horario TEXT NOT NULL,
                grupo TEXT NOT NULL,
                id_grupo TEXT NOT NULL,
                sucursal TEXT,
                pago REAL NOT NULL,
                ahorro REAL NOT NULL,
                mensaje_original TEXT NOT NULL,
                archivo_origen TEXT NOT NULL,
                remitente_whatsapp TEXT NOT NULL,
                hash_mensaje TEXT UNIQUE NOT NULL
            )
        ''')
        
        # Índices para mejorar consultas
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_fecha_mensaje 
            ON pagos(fecha_mensaje)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_corte_horario 
            ON pagos(corte_horario)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_grupo 
            ON pagos(grupo)
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_hash 
            ON pagos(hash_mensaje)
        ''')
        
        conn.commit()
        conn.close()
    
    def existe_pago(self, hash_mensaje: str) -> bool:
        """Verifica si un pago ya fue procesado (evita duplicados)"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute(
            'SELECT COUNT(*) FROM pagos WHERE hash_mensaje = ?',
            (hash_mensaje,)
        )
        
        count = cursor.fetchone()[0]
        conn.close()
        
        return count > 0
    
    def insertar_pago(self, pago: Dict, archivo_origen: str) -> bool:
        """Inserta un pago en la base de datos si no existe"""
        # Verificar si ya existe
        if self.existe_pago(pago['hash']):
            return False
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT INTO pagos (
                    fecha_mensaje,
                    fecha_procesamiento,
                    corte_horario,
                    grupo,
                    id_grupo,
                    sucursal,
                    pago,
                    ahorro,
                    mensaje_original,
                    archivo_origen,
                    remitente_whatsapp,
                    hash_mensaje
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                pago['fecha_hora'],
                datetime.now(),
                pago['corte_horario'],
                pago['grupo'],
                pago['id_grupo'],
                pago.get('sucursal', 'N/A'),
                pago['pago'],
                pago['ahorro'],
                pago['mensaje_original'],
                archivo_origen,
                pago['remitente'],
                pago['hash']
            ))
            
            conn.commit()
            conn.close()
            return True
            
        except sqlite3.IntegrityError:
            # Hash duplicado (otro proceso lo insertó)
            conn.close()
            return False
        except Exception as e:
            conn.close()
            raise e
    
    def insertar_pagos_lote(self, pagos: List[Dict], archivo_origen: str) -> Dict[str, int]:
        """Inserta múltiples pagos y retorna estadísticas"""
        insertados = 0
        duplicados = 0
        errores = 0
        
        for pago in pagos:
            try:
                if self.insertar_pago(pago, archivo_origen):
                    insertados += 1
                else:
                    duplicados += 1
            except Exception as e:
                errores += 1
                print(f"Error insertando pago: {e}")
        
        return {
            'insertados': insertados,
            'duplicados': duplicados,
            'errores': errores,
            'total': len(pagos)
        }
    
    def obtener_todos_pagos(self) -> List[Dict]:
        """Obtiene todos los pagos de la base de datos"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT 
                id,
                fecha_mensaje,
                fecha_procesamiento,
                corte_horario,
                grupo,
                id_grupo,
                sucursal,
                pago,
                ahorro,
                mensaje_original,
                archivo_origen,
                remitente_whatsapp
            FROM pagos
            ORDER BY fecha_mensaje DESC
        ''')
        
        pagos = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return pagos
    
    def obtener_pagos_por_fecha(self, fecha_inicio: str, fecha_fin: str) -> List[Dict]:
        """Obtiene pagos en un rango de fechas"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM pagos
            WHERE DATE(fecha_mensaje) BETWEEN ? AND ?
            ORDER BY fecha_mensaje DESC
        ''', (fecha_inicio, fecha_fin))
        
        pagos = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return pagos
    
    def obtener_pagos_por_corte(self, corte: str, fecha: Optional[str] = None) -> List[Dict]:
        """Obtiene pagos de un corte específico"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        if fecha:
            cursor.execute('''
                SELECT * FROM pagos
                WHERE corte_horario = ? AND DATE(fecha_mensaje) = ?
                ORDER BY fecha_mensaje DESC
            ''', (corte, fecha))
        else:
            cursor.execute('''
                SELECT * FROM pagos
                WHERE corte_horario = ?
                ORDER BY fecha_mensaje DESC
            ''', (corte,))
        
        pagos = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return pagos
    
    def obtener_estadisticas(self) -> Dict:
        """Obtiene estadísticas generales de la base de datos"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Total de pagos
        cursor.execute('SELECT COUNT(*) FROM pagos')
        total_pagos = cursor.fetchone()[0]
        
        # Suma total de pagos y ahorros
        cursor.execute('SELECT SUM(pago), SUM(ahorro) FROM pagos')
        suma_pago, suma_ahorro = cursor.fetchone()
        
        # Pagos por corte
        cursor.execute('''
            SELECT corte_horario, COUNT(*), SUM(pago), SUM(ahorro)
            FROM pagos
            GROUP BY corte_horario
        ''')
        por_corte = cursor.fetchall()
        
        # Pagos por grupo (top 10)
        cursor.execute('''
            SELECT grupo, COUNT(*), SUM(pago), SUM(ahorro)
            FROM pagos
            GROUP BY grupo
            ORDER BY SUM(pago) DESC
            LIMIT 10
        ''')
        por_grupo = cursor.fetchall()
        
        # Pagos por sucursal
        cursor.execute('''
            SELECT sucursal, COUNT(*), SUM(pago), SUM(ahorro)
            FROM pagos
            GROUP BY sucursal
            ORDER BY SUM(pago) DESC
        ''')
        por_sucursal = cursor.fetchall()
        
        conn.close()
        
        return {
            'total_pagos': total_pagos,
            'suma_total_pago': suma_pago or 0,
            'suma_total_ahorro': suma_ahorro or 0,
            'por_corte': por_corte,
            'por_grupo': por_grupo,
            'por_sucursal': por_sucursal
        }


if __name__ == "__main__":
    # Prueba básica
    db = DatabaseManager()
    print("Base de datos inicializada correctamente")
    print(f"Ruta: {db.db_path}")
    stats = db.obtener_estadisticas()
    print(f"Total de pagos en BD: {stats['total_pagos']}")

