from datetime import date, timedelta
import sqlite3
import re
import os, sys

def _base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

DB_NAME = os.path.join(_base_path(), "gastos_andres.db")
TIPOS_VALIDOS = ("Personal", "Negocio")
MONTO_MAXIMO   = 999_999_999
DESC_MAX_CHARS = 200

CATEGORIAS_DEFAULT = [
    "Compra de repuestos",
    "Almuerzo",
    "Transporte",
    "Arriendo local",
    "Internet",
    "Servicios públicos",
    "Nómina",
    "Varios",
]

def _parsear_monto(monto):
    monto_str = str(monto).strip()
    if not re.match(r'^[\d.,]+$', monto_str):
        raise ValueError(f"El monto '{monto}' debe ser un número válido.")
 
    tiene_punto = '.' in monto_str
    tiene_coma  = ',' in monto_str
 
    if tiene_punto and tiene_coma:
        if monto_str.rfind(',') > monto_str.rfind('.'):
            limpio = monto_str.replace('.', '').replace(',', '.')
        else:
            limpio = monto_str.replace(',', '')
    elif tiene_coma and not tiene_punto:
        partes = monto_str.split(',')
        if len(partes) > 2:
            limpio = monto_str.replace(',', '')
        else:
            limpio = monto_str.replace(',', '.')
    elif tiene_punto and not tiene_coma:
        partes = monto_str.split('.')
        if len(partes) > 2:
            limpio = monto_str.replace('.', '')
        elif len(partes[1]) <= 2:
            limpio = monto_str
        else:
            limpio = monto_str.replace('.', '')
    else:
        limpio = monto_str
    return float(limpio)
 
 
# Función para iniciar el programa con el archivo SQLite funcionando desde el inicio
def inicializar_db():
    try:
        with sqlite3.connect(DB_NAME) as conn:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS gastos (
                        id          INTEGER PRIMARY KEY AUTOINCREMENT,
                        fecha       TEXT    NOT NULL,
                        descripcion TEXT    NOT NULL,
                        categoria   TEXT    NOT NULL,
                        monto       REAL    NOT NULL,
                        tipo        TEXT    NOT NULL CHECK(tipo IN ('Personal', 'Negocio'))
                    )
            """)
            conn.commit()
        return True, "Base de datos inicializada correctamente."
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"
 
# Función para Agregar los gastos al programa
def agregar_gasto(descripcion, categoria, monto, tipo):
    """
    Inserta un gasto nuevo en la base de datos.
    La fecha se asigna automáticamente con la fecha de hoy.
    """
    if not isinstance(descripcion, str):
        return False, "La descripción debe ser un texto."
    descripcion = descripcion.strip()
    if not descripcion:
        return False, "La Descripción no puede estar vacía."
    if len(descripcion) > DESC_MAX_CHARS:
        return False, f"La descripción no puede superar los {DESC_MAX_CHARS} caracteres."
    
    if not isinstance(categoria, str):
        return False, "La categoría debe ser un texto"
    categoria = categoria.strip()
    if not categoria:
        return False, "La categoría no puede estar vacía."
    if len(categoria) > DESC_MAX_CHARS:
        return False, f"La categoría no puede superar los {DESC_MAX_CHARS} caracteres."
    
    if not isinstance(tipo, str):
        return False, "El tipo debe ser un texto."
    tipo = tipo.strip().capitalize()
    if tipo not in TIPOS_VALIDOS:
        return False, f"Tipo inválido: '{tipo}'. Debe ser 'Personal' o 'Negocio'."
    
    try:
        monto_limpio = _parsear_monto(monto)
    except ValueError as e:
        return False, str(e)
    if monto_limpio <= 0:
        return False, f"El monto debe ser un número mayor a 0."
    if monto_limpio > MONTO_MAXIMO:
        return False, f"El monto supera el maximo permitido de ${MONTO_MAXIMO:,}."
 
    fecha_hoy = date.today().isoformat()
    try:
        with sqlite3.connect(DB_NAME) as conn:
            conn.execute("""
            INSERT INTO gastos (fecha, descripcion, categoria, monto, tipo)
            VALUES (?, ?, ?, ?, ?)
            """, (fecha_hoy, descripcion, categoria, monto_limpio, tipo))
            conn.commit()
        return True, "Gasto registrado correctamente."
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"
 
# Función para obtener los gastos del dia
def obtener_gastos_dia(fecha=None):
    """
    Devuelve todos los gastos registrados de un día específico.
    Si no se pasa la fecha, usa la de hoy automáticamente.
    Recuerda que siempre que necesitas datos de retorno, usas cursor = conn.cursor(), porque
    es la función que te permitirá obtener los datos que quieras retornar.
    """
    if fecha is None:
        fecha = date.today().isoformat()
    else:
        if not isinstance(fecha, str):
            return False, "La fecha debe ser un texto en formato YYYY-MM-DD."
        try:
            date.fromisoformat(fecha)
        except ValueError:
            return False, f"La fecha '{fecha}' no es valida, usa el formato YYYY-MM-DD."
        
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, fecha, descripcion, categoria, monto, tipo
                FROM gastos
                WHERE fecha = ?
                ORDER BY id ASC
            """, (fecha,))
            filas = cursor.fetchall()
            return True, {
                "data": filas,
                "count": len(filas)
            }
 
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"
 
# Función para obtener los gastos del mes
def obtener_gastos_mes(mes, año):
    """
    Devuelve todos los gastos de un mes y año específico.
    mes  → número entero, ej: 1 para enero
    año  → número entero, ej: 2024
    """
    if isinstance(mes, bool) or isinstance(año, bool):
        return False, "El mes y el año deben ser números enteros."
    if not isinstance(mes, int) or not isinstance(año, int):
        return False, "El mes y el año deben ser números enteros."
    if mes < 1 or mes > 12:
        return False, f"Mes inválido, ({mes}). Debe estar entre 1 y 12."
    if año < 2000 or año > 2060:
        return False, f"Año inválido, ({año}). Debe estar entre el 2000 y el 2060."
    
    mes_str = str(mes).zfill(2)
    patron = f"{año}-{mes_str}-%"
 
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, fecha, descripcion, categoria, monto, tipo
                FROM gastos
                WHERE fecha LIKE ?
                ORDER BY fecha ASC, id ASC 
            """, (patron,))
            filas = cursor.fetchall()
        return True, {
                "data": filas,
                "count": len(filas)
            }
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"
 
def obtener_por_categoria(mes, año):
    if isinstance(mes, bool) or isinstance(año, bool):
        return False, "El mes y el año deben ser números enteros."
    if not isinstance(mes, int) or not isinstance(año, int):
        return False, "El mes y el año deben ser números enteros."
    if mes < 1 or mes > 12:
        return False, f"Mes inválido, ({mes}). Debe estar entre 1 y 12."
    if año < 2000 or año > 2060:
        return False, f"Año inválido, ({año}). Debe estar entre el 2000 y el 2060."
    
    mes_str = str(mes).zfill(2)
    patron = f"{año}-{mes_str}-%"
 
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT categoria, SUM(monto) as total
                FROM gastos
                WHERE fecha LIKE ?
                GROUP BY categoria  
                ORDER BY total DESC
            """, (patron,))
            filas = cursor.fetchall()
        return True, {
                "data": filas,
                "count": len(filas)
            }
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"
 
def editar_gasto(id_gasto, fecha, categoria, descripcion, monto, tipo):
    if isinstance(id_gasto, bool) or not isinstance(id_gasto, int) or id_gasto <= 0:
        return False, "El ID debe ser un entero positivo."
 
    if not isinstance(fecha, str):
        return False, "La fecha debe estar en formato YYYY-MM-DD."
    try:
        date.fromisoformat(fecha)
    except ValueError:
        return False, f"La fecha '{fecha}' no es válida, usa el formato YYYY-MM-DD."
    
    if not isinstance(descripcion, str):
        return False, "La descripción debe ser un texto."
    descripcion = descripcion.strip()
    if not descripcion:
        return False, "La Descripción no puede estar vacía."
    if len(descripcion) > DESC_MAX_CHARS:
        return False, f"La descripción no puede superar los {DESC_MAX_CHARS} caracteres."
    
    if not isinstance(categoria, str):
        return False, "La categoría debe ser un texto"
    categoria = categoria.strip()
    if not categoria:
        return False, "La categoría no puede estar vacía."
    if len(categoria) > DESC_MAX_CHARS:
        return False, f"La categoría no puede superar los {DESC_MAX_CHARS} caracteres."
    if not isinstance(tipo, str):
        return False, "El tipo debe ser un texto."
    tipo = tipo.strip().capitalize()
    if tipo not in TIPOS_VALIDOS:
        return False, f"Tipo inválido: '{tipo}'. Debe ser 'Personal' o 'Negocio'."
    
    try:
        monto_limpio = _parsear_monto(monto)
    except ValueError as e:
        return False, str(e)
    if monto_limpio <= 0:
        return False, "El monto debe ser un número mayor a 0."
    if monto_limpio > MONTO_MAXIMO:
        return False, f"El monto supera el maximo permitido de ${MONTO_MAXIMO:,}"
 
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE gastos
                SET fecha = ?,
                categoria = ?,
                descripcion = ?,
                monto = ?,
                tipo = ?
                WHERE id = ?
            """, (fecha, categoria, descripcion, monto_limpio, tipo, id_gasto))
            conn.commit()
            if cursor.rowcount == 0:
                return False, f"No se encontró ningun gasto con esa ID {id_gasto}"
            
            return True, "Gasto actualizado correctamente."
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"
 
def eliminar_gasto(id_gasto):
    if isinstance(id_gasto, bool) or not isinstance(id_gasto, int) or id_gasto <= 0:
        return False, "El ID debe ser un entero positivo."
    try:
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                DELETE FROM gastos
                WHERE id = ?
            """, (id_gasto,))
            conn.commit()
            if cursor.rowcount == 0:
                return False, f"No se encontró ningun gasto con esa ID {id_gasto}"
            return True, "Gasto eliminado correctamente."
    except sqlite3.Error as e:
        return False, f"Error en la base de datos: {e}"
    except Exception as e:
        return False, f"Error inesperado: {e}"











