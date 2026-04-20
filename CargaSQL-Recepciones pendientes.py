import psycopg2
import pandas as pd
from datetime import datetime, timedelta  # ← importamos timedelta
from sqlalchemy import create_engine

# Configuración de conexión a PostgreSQL
DB_CONFIG = {
    'host': 'User-host',
    'port': xxxx,
    'database': 'Retail_MARKET',
    'user': 'USER_BD',
    'password': 'xxxxx'
}

# Query para extraer los datos
QUERY_EXTRACT = """
    SELECT 
        numero AS "Número",
        documento_asociado AS "Número de orden",
        proveedor AS "Proveedor",
        sucursal_recepcion AS "Sucursal",
        estatus_compra AS "Estatus compra",
        estatus AS "Estatus ODC",
        factura AS "Factura",
        fecha_de_creacion AS "F. Crea. OC",
        fecha_de_validacion AS "F. Valid.",
        receptor_responsable AS "Receptor",
        validada_por AS "Validada por",
        cantidad_de_items AS "Cant. ítems",
        area_de_recepcion AS "Área de recepción",
        monto_total_usd AS "Monto total $"
    FROM public.recepcion
    WHERE tipo_de_recepcion = 'Recepcion-Validación' 
      AND estatus_compra = 'Pendiente por validar'
"""

# Nombre de la tabla destino
TABLE_DESTINO = 'recepciones_pendientes'

def eliminar_registros_fecha_anterior(conn):
    """Elimina los registros de la tabla destino que corresponden a la fecha anterior (día actual - 1)"""
    fecha_anterior = (datetime.now() - timedelta(days=1)).date()   # ← fecha del día anterior
    
    # Query para eliminar registros de esa fecha
    delete_query = """
        DELETE FROM public.{}
        WHERE DATE(fecha_carga) = %s
    """.format(TABLE_DESTINO)
    
    try:
        cursor = conn.cursor()
        cursor.execute(delete_query, (fecha_anterior,))
        registros_eliminados = cursor.rowcount
        conn.commit()
        cursor.close()
        
        if registros_eliminados > 0:
            print(f"✓ Se eliminaron {registros_eliminados} registros de la fecha {fecha_anterior}")
        else:
            print(f"No existían registros previos para la fecha {fecha_anterior}")
        
        return registros_eliminados
    except Exception as e:
        print(f"❌ Error al eliminar registros: {e}")
        conn.rollback()
        raise

def agregar_fecha_carga(df):
    """Agrega la columna fecha_carga con la fecha del día anterior a cada registro"""
    fecha_anterior = datetime.now() - timedelta(days=1)   # ← fecha del día anterior
    df['fecha_carga'] = fecha_anterior
    return df

def main():
    try:
        # 1. Establecer conexión
        print("Conectando a la base de datos...")
        conn = psycopg2.connect(**DB_CONFIG)
        
        # 2. Extraer datos a DataFrame
        print("Ejecutando consulta y extrayendo datos...")
        df = pd.read_sql_query(QUERY_EXTRACT, conn)
        
        # Renombrar columnas para que coincidan con la tabla destino
        df = df.rename(columns={
            "Número": "numero",
            "Número de orden": "numero_de_orden",
            "Proveedor": "proveedor",
            "Factura": "factura",
            "Sucursal": "sucursal",
            "Estatus ODC": "estatus_odc",
            "Estatus compra": "estatus_compra",
            "F. Crea. OC": "f_crea_oc",
            "F. Valid.": "f_valid",
            "Receptor": "receptor",
            "Validada por": "validada_por",
            "Cant. ítems": "cant_items",
            "Área de recepción": "area_de_recepcion",
            "Monto total $": "monto_total_usd"
        })
        
        print(f"Se extrajeron {len(df)} registros de la consulta")
        
        # 3. Agregar columna de fecha_carga (día anterior)
        df = agregar_fecha_carga(df)
        
        # 4. Eliminar registros existentes de la fecha anterior
        print("\n--- Limpiando registros anteriores del día anterior ---")
        eliminar_registros_fecha_anterior(conn)   # ← nombre actualizado
        
        # 5. Cargar datos a la tabla destino
        if not df.empty:
            print(f"\n--- Cargando nuevos datos a la tabla '{TABLE_DESTINO}' ---")
            
            # Crear engine de SQLAlchemy para pandas
            engine = create_engine(
                f"postgresql://{DB_CONFIG['user']}:{DB_CONFIG['password']}@"
                f"{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
            )
            
            # Cargar datos usando 'append' (la tabla ya fue limpiada para la fecha anterior)
            df.to_sql(
                TABLE_DESTINO,
                engine,
                if_exists='append',
                index=False,
                method='multi'
            )
            
            fecha_carga_str = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S')
            print(f"✓ ¡Carga completada! Se insertaron {len(df)} registros en '{TABLE_DESTINO}'")
            print(f"  Fecha de carga asignada: {fecha_carga_str}")
        else:
            print("\nNo hay registros nuevos para insertar")
        
        # Cerrar conexión
        conn.close()
        
    except Exception as e:
        print(f"❌ Error general: {e}")
        if isinstance(e, psycopg2.Error):
            print(f"Detalle PostgreSQL: {e.pgerror}")

if __name__ == "__main__":
    main()