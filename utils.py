import os
import pandas as pd
import logging
from conn import connection

def setup_logging(bot_name, log_file=None):
    """Configurar logging para un bot específico"""
    if log_file is None:
        log_file = f"{bot_name}_bot.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(bot_name)

def ensure_directories():
    """Asegura que todas las carpetas necesarias existan"""
    directories = ['no_procesados', 'Errores', 'Temp', 'Recetas']
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

def consultarCadenaFrio(oc_numero: str) -> bool:
    """
    Devuelve True si la orden de compra es de cadena de frío, False si es seco.

    Parámetros:
    - oc_numero: str. Número de orden de compra.

    Retorna:
    - bool: True si es frío, False si es seco.
    """
    conn = connection('PRD')
    query = f"""
        SELECT m.ZZCADENA_FRIO, m.MATNR
        FROM EKPO e
        JOIN MARA m ON m.MATNR = e.MATNR 
        WHERE e.EBELN = '{oc_numero}'
    """
    df_oc = pd.read_sql_query(query, conn)

    # Verifica si al menos un valor en la columna es 'X'
    if (df_oc["ZZCADENA_FRIO"] == 'X').any():
        return True
    else:
        return False
    

def devolverEanOC(oc_numero):
    conn = connection('PRD')
    query = f"""
        SELECT DISTINCT e.EAN11, e.MENGE
        FROM MARA m
        JOIN EKPO e ON m.MANDT = e.MANDT
        WHERE e.EBELN = '{oc_numero}'
        """
    df_ean = pd.read_sql_query(query, conn)
    return df_ean

def obtener_mapping_ean_material(oc_numero):
    """
    Obtiene el mapeo entre EAN y código de material para una OC.
    
    Args:
        oc_numero: Número de orden de compra
        
    Returns:
        dict: Diccionario {EAN: MATNR}
    """
    conn = connection('PRD')
    query = f"""
        SELECT DISTINCT 
            m.EAN11,
            e.MATNR
        FROM EKPO e
        JOIN MARA m ON e.MATNR = m.MATNR AND e.MANDT = m.MANDT
        WHERE e.EBELN = '{oc_numero}'
        AND m.EAN11 IS NOT NULL
    """
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    return dict(zip(df['EAN11'], df['MATNR']))


def validar_estructura_excel(df_excel):
    """
    Valida que el Excel tenga las columnas requeridas.
    
    Args:
        df_excel: DataFrame del Excel
        
    Returns:
        tuple: (es_valido, mensaje_error)
    """
    columnas_requeridas = [
        'Remito y Nro. Entrega',
        'Cant confirmada',
        'Lote estuche',
        'Fecha Vencimiento'
    ]
    
    columnas_faltantes = [col for col in columnas_requeridas if col not in df_excel.columns]
    
    if columnas_faltantes:
        return False, f"Columnas faltantes en Excel: {', '.join(columnas_faltantes)}"
    
    return True, ""


def generar_reporte_consolidacion(oc_numero, eans_consolidados, path_salida):
    """
    Genera un reporte de la consolidación realizada.
    
    Args:
        oc_numero: Número de OC
        eans_consolidados: Lista de diccionarios con info de consolidación
        path_salida: Ruta donde guardar el reporte
    """
    import datetime
    
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"consolidacion_oc_{oc_numero}_{timestamp}.txt"
    filepath = os.path.join(path_salida, filename)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(f"REPORTE DE CONSOLIDACIÓN DE CANTIDADES\n")
        f.write(f"{'=' * 50}\n")
        f.write(f"OC: {oc_numero}\n")
        f.write(f"Fecha: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"{'=' * 50}\n\n")
        
        for ean_info in eans_consolidados:
            f.write(f"EAN: {ean_info['ean']}\n")
            f.write(f"  Filas originales: {ean_info['filas_originales']}\n")
            f.write(f"  Cantidad total: {ean_info['cantidad_total']}\n")
            f.write(f"  Cantidad pendiente SAP: {ean_info['cantidad_pendiente']}\n")
            f.write(f"  Estado: {ean_info['estado']}\n")
            f.write(f"{'-' * 30}\n")
    
    return filepath