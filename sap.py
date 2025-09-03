import os
import time
import logging
import pandas as pd
import win32com.client
from dotenv import load_dotenv
from abrirsap import ingresarsap
from utils import consultarCadenaFrio
import pythoncom
import shutil
from datetime import datetime
# Configuración de logging
default_format = "%(asctime)s [%(levelname)s] %(message)s"
logging.basicConfig(level=logging.INFO, format=default_format, datefmt="%Y-%m-%d %H:%M:%S")
logger = logging.getLogger(__name__)

# Carga de variables de entorno
try:
    load_dotenv()
    logger.info("Variables de entorno cargadas correctamente.")
except Exception as e:
    logger.warning("No se pudo cargar .env: %s", e)



def get_sap_session(sesionsap: int = 0):
    """
    Inicializa el entorno COM y obtiene la sesión SAP.
    
    Args:
        sesionsap: Índice de la sesión SAP a utilizar
        
    Returns:
        CDispatch: Objeto de sesión SAP o None si falla
    """
    pythoncom.CoInitialize()
    
    try:
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
            if not isinstance(SapGuiAuto, win32com.client.CDispatch):
                print("No se pudo obtener el objeto SAPGUI. Asegúrate de que SAP GUI está abierto.")
                return None
        except Exception as e:
            print(f"Error al conectar con SAP GUI: {e}")
            print("Asegúrate que SAP GUI está abierto y que el scripting está habilitado.")
            return None
            
        try:
            application = SapGuiAuto.GetScriptingEngine
            if not isinstance(application, win32com.client.CDispatch):
                print("No se pudo obtener el ScriptingEngine.")
                return None
        except Exception as e:
            print(f"Error al obtener ScriptingEngine: {e}")
            print("Es posible que el scripting esté deshabilitado en SAP GUI.")
            return None
            
        try:
            if application.Children.Count == 0:
                print("No hay conexiones SAP activas. Inicia sesión en SAP primero.")
                return None
            connection = application.Children(0)
            if not isinstance(connection, win32com.client.CDispatch):
                print("No se pudo obtener la conexión SAP.")
                return None
        except Exception as e:
            print(f"Error al acceder a las conexiones SAP: {e}")
            return None
            
        try:
            session_count = connection.Children.Count
            print(f"Sesiones disponibles: {session_count}")
            if session_count == 0:
                print("No hay sesiones disponibles en la conexión")
                return None
            if sesionsap >= session_count:
                print(f"Índice de sesión {sesionsap} no válido. Solo hay {session_count} sesiones.")
                sesionsap = 0
            session = connection.Children(sesionsap)
        except Exception as e:
            print(f"Error al acceder a las sesiones: {e}")
            return None
            
        try:
            session_type = session.Type
            print(f"Sesión tipo: {session_type} obtenida correctamente")
            if hasattr(session, 'Info'):
                try:
                    if hasattr(session.Info, 'ScreenName'):
                        status = session.Info.ScreenName
                        print(f"Pantalla actual: {status}")
                    else:
                        print("La propiedad ScreenName no está disponible")
                except:
                    print("No se pudo acceder a ScreenName, pero la sesión parece válida")
            else:
                print("El objeto Info no está disponible, pero la sesión parece válida")
            return session
        except Exception as e:
            print(f"La sesión no está operativa: {e}")
            return None
    except Exception as e:
        print(f"Error no manejado al obtener la sesión SAP: {e}")
        return None



def normalize_sap_number(value):
    """
    Normaliza valores numéricos devueltos por SAP que pueden venir en formato científico o decimal.
    Maneja números con formato europeo (punto como separador de miles).
    
    Args:
        value: Valor devuelto por SAP (puede ser string, float, int, etc.)
        
    Returns:
        str: Valor normalizado como string entero
    """
    try:
        # Convertir a string primero
        str_value = str(value).strip()
        
        # Si está vacío, retornar "0"
        if not str_value:
            return "0"

        # Si es un número en formato científico (ej: "1.0E+3")
        if 'E' in str_value.upper() or 'e' in str_value:
            try:
                float_val = float(str_value)
                return str(int(float_val))
            except:
                return str_value
        
        # Manejar formato europeo (punto como separador de miles)
        # Ejemplo: "1.000" -> 1000, "1.500" -> 1500
        if '.' in str_value and ',' not in str_value:
            # Verificar si es formato europeo (punto como separador de miles)
            # Si tiene más de 3 dígitos después del punto, es formato europeo
            parts = str_value.split('.')
            if len(parts) == 2 and len(parts[1]) == 3:
                # Es formato europeo: "1.000" -> 1000
                try:
                    # Remover el punto y convertir a entero
                    clean_value = str_value.replace('.', '')
                    int_val = int(clean_value)
                    return str(int_val)
                except:
                    pass
        
        # Si es un float con decimales (ej: "1000.0")
        if '.' in str_value:
            try:
                float_val = float(str_value)
                return str(int(float_val))
            except:
                return str_value
        
        # Si es un entero normal
        try:
            int_val = int(float(str_value))
            return str(int_val)
        except:
            return str_value
            
    except Exception as e:
        logger.warning(f"Error normalizando valor SAP '{value}': {e}")
        return str(value)


def find_row_by_ean(grid, ean_to_find, start_index=0):
    """
    Recorre el grid buscando el EAN en la columna ZZEAN13.
    Devuelve el índice de fila donde coincide, o None si no lo encuentra.
    """
    total = grid.RowCount
    for offset in range(total):
        idx = (start_index + offset) % total
        try:
            current = grid.getCellValue(idx, "ZZEAN13").strip()
        except Exception:
            continue
        if current == ean_to_find.strip():
            return idx
    return None


def get_quantity_column_name(grid):
    """
    Determina el nombre correcto de la columna de cantidad en el grid de SAP.
    
    Args:
        grid: Grid de SAP
        
    Returns:
        str: Nombre de la columna de cantidad
    """
    # Lista de posibles nombres de columnas de cantidad
    possible_names = ["MENGE", "CANTIDAD", "QUANTITY", "QTY", "AMOUNT"]
    
    # Intentar obtener el nombre de la primera fila
    try:
        for col_name in possible_names:
            try:
                # Intentar acceder a la columna
                test_value = grid.getCellValue(0, col_name)
                logger.info(f"Columna de cantidad encontrada: {col_name}")
                return col_name
            except:
                continue
    except:
        pass
    
    # Si no se encuentra, usar el nombre por defecto
    logger.warning("No se pudo determinar la columna de cantidad, usando 'MENGE' por defecto")
    return "MENGE"


def debug_grid_columns(grid):
    """
    Función de debug para listar todas las columnas disponibles en el grid.
    """
    try:
        logger.info("🔍 DEBUG: Listando columnas disponibles en el grid...")
        # Intentar obtener información de las primeras filas para detectar columnas
        for row_idx in range(min(3, grid.RowCount)):
            logger.info(f"Fila {row_idx}:")
            for col_idx in range(20):  # Probar hasta 20 columnas
                try:
                    col_name = f"Col{col_idx}"
                    value = grid.getCellValue(row_idx, col_name)
                    if value:
                        logger.info(f"  {col_name}: {value}")
                except:
                    continue
    except Exception as e:
        logger.error(f"Error en debug_grid_columns: {e}")


def find_row_by_ean_and_quantity(grid, ean_to_find, expected_quantity, start_index=0):
    """
    Recorre el grid buscando el EAN en la columna ZZEAN13 y valida la cantidad pendiente.
    Devuelve el índice de fila donde coincide, o None si no lo encuentra.
    
    Args:
        grid: Grid de SAP
        ean_to_find: EAN a buscar
        expected_quantity: Cantidad esperada (del Excel)
        start_index: Índice inicial para la búsqueda
        
    Returns:
        int: Índice de fila donde coincide, o None si no lo encuentra
    """
    # Determinar el nombre de la columna de cantidad
    quantity_column = get_quantity_column_name(grid)
    
    total = grid.RowCount
    for offset in range(total):
        idx = (start_index + offset) % total
        try:
            # Obtener EAN de la fila
            current_ean = grid.getCellValue(idx, "ZZEAN13").strip()
            
            # Si el EAN coincide, validar cantidad pendiente
            if current_ean == ean_to_find.strip():
                # Obtener cantidad pendiente de SAP usando la columna correcta
                sap_quantity = grid.getCellValue(idx, "CANT_PEND")
                sap_quantity_normalized = normalize_sap_number(sap_quantity)
                expected_quantity_str = str(int(expected_quantity))
                
                logger.info(f"EAN encontrado: {current_ean}")
                logger.info(f"Cantidad SAP (original): {sap_quantity}")
                logger.info(f"Cantidad SAP (normalizada): {sap_quantity_normalized}")
                logger.info(f"Cantidad esperada (Excel): {expected_quantity_str}")
                
                # Comparar cantidades normalizadas
                if sap_quantity_normalized == expected_quantity_str:
                    logger.info(f"✅ Cantidad pendiente válida: {sap_quantity_normalized}")
                    return idx
                else:
                    logger.warning(f"❌ Cantidad pendiente no coincide: SAP={sap_quantity_normalized}, Excel={expected_quantity_str}")
                    # Continuar buscando en caso de que haya otra fila con el mismo EAN
                    continue
                    
        except Exception as e:
            logger.warning(f"Error accediendo a fila {idx}: {e}")
            continue
    
    return None


def registrar_error_ean_no_encontrado(oc, ean_no_encontrado, path_excel):
    """
    Registra el error de EAN no encontrado en un archivo de log específico.
    
    Args:
        oc: Número de orden de compra
        ean_no_encontrado: EAN que no se encontró en SAP
        path_excel: Ruta del archivo Excel
    """
    try:
        import os
        from datetime import datetime
        
        # Crear directorio de errores si no existe
        error_dir = os.path.join(os.getcwd(), "Errores")
        os.makedirs(error_dir, exist_ok=True)
        
        # Crear archivo de error específico para EAN no encontrado
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        error_file = os.path.join(error_dir, f"error_ean_no_encontrado_{oc}_{timestamp}.txt")
        
        with open(error_file, "w", encoding="utf-8") as f:
            f.write("ERROR DE EAN NO ENCONTRADO - OC CANCELADA\n")
            f.write("=" * 50 + "\n")
            f.write(f"OC: {oc}\n")
            f.write(f"Fecha: {datetime.now()}\n")
            f.write(f"EAN no encontrado: {ean_no_encontrado}\n")
            f.write(f"Archivo Excel: {path_excel}\n")
            f.write(f"Motivo: EAN del Excel no existe en ninguna fila de SAP\n")
            f.write(f"Estado: OC cancelada - flujo cortado\n")
            f.write("=" * 50 + "\n")
            f.write("\nEl sistema requiere que todos los EANs del Excel existan en SAP.\n")
            f.write("Esta OC ha sido cancelada debido a EANs no encontrados.\n")
            f.write("Se procede con la siguiente orden de compra.\n")
        
        logger.info(f"📝 Error de EAN no encontrado registrado en: {error_file}")
        
    except Exception as e:
        logger.error(f"❌ Error registrando error de EAN no encontrado para OC {oc}: {e}")


def registrar_error_ean_repetido(oc, ean_repetido, motivo, path_excel):
    """
    Registra el error de EAN repetido en un archivo de log específico.
    
    Args:
        oc: Número de orden de compra
        ean_repetido: EAN que causó el error
        motivo: Motivo del error
        path_excel: Ruta del archivo Excel
    """
    try:
        import os
        from datetime import datetime
        
        # Crear directorio de errores si no existe
        error_dir = os.path.join(os.getcwd(), "Errores")
        os.makedirs(error_dir, exist_ok=True)
        
        # Crear archivo de error específico para EAN repetido
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        error_file = os.path.join(error_dir, f"error_ean_repetido_{oc}_{timestamp}.txt")
        
        with open(error_file, "w", encoding="utf-8") as f:
            f.write("ERROR DE EAN REPETIDO - OC CANCELADA\n")
            f.write("=" * 50 + "\n")
            f.write(f"OC: {oc}\n")
            f.write(f"Fecha: {datetime.now()}\n")
            f.write(f"EAN repetido: {ean_repetido}\n")
            f.write(f"Archivo Excel: {path_excel}\n")
            f.write(f"Motivo: {motivo}\n")
            f.write(f"Estado: OC cancelada - flujo cortado\n")
            f.write("=" * 50 + "\n")
            f.write("\nEl sistema requiere que los EANs repetidos cumplan con las validaciones.\n")
            f.write("Esta OC ha sido cancelada debido a problemas con EANs repetidos.\n")
            f.write("Se procede con la siguiente orden de compra.\n")
        
        logger.info(f"📝 Error de EAN repetido registrado en: {error_file}")
        
    except Exception as e:
        logger.error(f"❌ Error registrando error de EAN repetido para OC {oc}: {e}")


def find_best_sap_row_for_ean(grid, ean_to_find, expected_quantity):
    """
    Busca la mejor fila de SAP para un EAN específico, considerando cantidad pendiente.
    
    Args:
        grid: Grid de SAP
        ean_to_find: EAN a buscar
        expected_quantity: Cantidad esperada del Excel
        
    Returns:
        tuple: (fila_encontrada, cantidad_sap, mensaje) o (None, None, mensaje_error)
    """
    try:
        total_rows = grid.RowCount
        filas_coincidentes = []
        
        # Buscar todas las filas que coincidan con el EAN
        logger.info(f"🔍 Buscando EAN '{ean_to_find}' en {total_rows} filas de SAP...")
        for sap_idx in range(total_rows):
            try:
                ean_sap = grid.getCellValue(sap_idx, "ZZEAN13").strip()
                logger.info(f"   - Fila {sap_idx}: EAN SAP='{ean_sap}' vs EAN buscado='{ean_to_find}'")
                if ean_sap == ean_to_find.strip():
                    logger.info(f"     ✅ EAN encontrado en fila {sap_idx}")
                    # Obtener cantidad pendiente de SAP usando el campo correcto
                    sap_quantity = grid.getCellValue(sap_idx, "CANT_PEND")
                    sap_quantity_normalized = normalize_sap_number(sap_quantity)
                    
                    filas_coincidentes.append({
                        'fila': sap_idx,
                        'cantidad_sap': sap_quantity_normalized,
                        'cantidad_original': sap_quantity
                    })
                    
            except Exception as e:
                logger.warning(f"Error accediendo a fila SAP {sap_idx}: {e}")
                continue
        
        if not filas_coincidentes:
            return None, None, f"EAN '{ean_to_find}' no encontrado en ninguna fila de SAP"
        
        # Si solo hay una fila, usarla
        if len(filas_coincidentes) == 1:
            fila = filas_coincidentes[0]
            return fila['fila'], fila['cantidad_sap'], f"EAN encontrado en fila {fila['fila']}"
        
        # Si hay múltiples filas, buscar la que coincida con la cantidad
        expected_quantity_str = str(int(expected_quantity))
        for fila in filas_coincidentes:
            if fila['cantidad_sap'] == expected_quantity_str:
                return fila['fila'], fila['cantidad_sap'], f"EAN y cantidad coinciden en fila {fila['fila']}"
        
        # Si ninguna coincide con la cantidad, usar la primera y registrar advertencia
        primera_fila = filas_coincidentes[0]
        logger.warning(f"⚠️ Múltiples filas para EAN '{ean_to_find}'. Usando primera fila {primera_fila['fila']}")
        return primera_fila['fila'], primera_fila['cantidad_sap'], f"Múltiples filas encontradas, usando primera"
        
    except Exception as e:
        logger.error(f"Error en find_best_sap_row_for_ean: {e}")
        return None, None, f"Error interno: {e}"


def validar_cantidades_ean_repetido(grid, ean, total_cantidad_excel):
    """
    Valida que la suma de cantidades de un EAN repetido no exceda la cantidad solicitada en SAP.
    
    Args:
        grid: Grid de SAP
        ean: EAN a validar
        total_cantidad_excel: Suma total de cantidades del Excel para este EAN
        
    Returns:
        tuple: (es_valido, cantidad_sap, mensaje)
    """
    try:
        # Buscar todas las filas de SAP que contengan este EAN
        filas_sap_ean = []
        total_rows = grid.RowCount
        
        logger.info(f"🔍 Buscando EAN {ean} en {total_rows} filas de SAP")
        
        # Debug: Listar columnas disponibles
        debug_grid_columns(grid)
        
        for sap_idx in range(total_rows):
            try:
                # Buscar EAN en la columna original
                ean_sap = grid.getCellValue(sap_idx, "ZZEAN13").strip()
                
                if ean_sap == ean.strip():
                    # Obtener cantidad pendiente de SAP usando el campo correcto
                    cantidad_sap = grid.getCellValue(sap_idx, "CANT_PEND")
                    
                    logger.info(f"📊 Fila {sap_idx}: EAN={ean_sap}, Campo cantidad=0,CANT_PEND, Valor={cantidad_sap}")
                    
                    if cantidad_sap:
                        cantidad_sap_normalizada = normalize_sap_number(cantidad_sap)
                        filas_sap_ean.append({
                            'fila': sap_idx,
                            'cantidad': int(cantidad_sap_normalizada),
                            'cantidad_original': cantidad_sap,
                            'columna_cantidad': "CANT_PEND"
                        })
                        logger.info(f"✅ EAN {ean} encontrado en fila {sap_idx}, cantidad: {cantidad_sap} -> {cantidad_sap_normalizada}")
                    else:
                        logger.warning(f"⚠️ EAN {ean} encontrado en fila {sap_idx} pero cantidad vacía")
                    
            except Exception as e:
                logger.warning(f"Error accediendo a fila SAP {sap_idx}: {e}")
                continue
        
        if not filas_sap_ean:
            return False, 0, f"EAN {ean} no encontrado en SAP"
        
        # Calcular cantidad total solicitada en SAP
        cantidad_total_sap = sum(fila['cantidad'] for fila in filas_sap_ean)
        
        logger.info(f"📊 Validación EAN {ean}:")
        logger.info(f"  - Cantidad Excel: {total_cantidad_excel}")
        logger.info(f"  - Cantidad SAP total: {cantidad_total_sap}")
        logger.info(f"  - Filas SAP encontradas: {len(filas_sap_ean)}")
        for fila in filas_sap_ean:
            logger.info(f"    - Fila {fila['fila']}: {fila['cantidad_original']} -> {fila['cantidad']} (col: {fila['columna_cantidad']})")
        
        # Validar que la cantidad del Excel no exceda la solicitada
        if total_cantidad_excel > cantidad_total_sap:
            return False, cantidad_total_sap, f"Cantidad Excel ({total_cantidad_excel}) excede cantidad SAP ({cantidad_total_sap})"
        elif total_cantidad_excel == cantidad_total_sap:
            return True, cantidad_total_sap, f"Cantidades coinciden: Excel={total_cantidad_excel}, SAP={cantidad_total_sap}"
        else:
            return True, cantidad_total_sap, f"Cantidad Excel ({total_cantidad_excel}) menor a SAP ({cantidad_total_sap}) - OK"
            
    except Exception as e:
        logger.error(f"Error validando cantidades para EAN {ean}: {e}")
        return False, 0, f"Error en validación: {e}"


def detectar_eans_repetidos_en_excel(df_excel):
    """
    Detecta EANs que aparecen en múltiples filas del Excel.
    
    Args:
        df_excel: DataFrame del Excel
        
    Returns:
        dict: Diccionario con EANs repetidos y sus filas correspondientes
    """
    try:
        eans_repetidos = {}
        
        # Agrupar por EAN y obtener todas las filas para cada EAN
        for ean, grupo in df_excel.groupby('EAN'):
            ean_str = str(ean).strip()
            if len(grupo) > 1:
                eans_repetidos[ean_str] = {
                    'filas': grupo.index.tolist(),
                    'cantidades': grupo['Cant confirmada'].tolist(),
                    'lotes': grupo['Lote estuche'].tolist(),
                    'fechas_vencimiento': grupo['Fecha Vencimiento'].tolist(),
                    'total_cantidad': grupo['Cant confirmada'].sum()
                }
        
        if eans_repetidos:
            logger.info(f"🔍 EANs repetidos detectados: {list(eans_repetidos.keys())}")
            for ean, info in eans_repetidos.items():
                logger.info(f"   - EAN {ean}: {len(info['filas'])} filas, total cantidad: {info['total_cantidad']}")
        
        return eans_repetidos
        
    except Exception as e:
        logger.error(f"Error detectando EANs repetidos: {e}")
        return {}


def agregar_fila_sap(grid, session, fila_actual):
    """
    Agrega una nueva fila en el grid de SAP usando el botón de agregar lote.
    
    Args:
        grid: Grid de SAP
        session: Sesión de SAP
        fila_actual: Índice de la fila actual donde estoy parado
        
    Returns:
        tuple: (True, nueva_fila_index) si se agregó exitosamente, (False, None) en caso contrario
    """
    try:
        logger.info(f"➕ Agregando nueva fila en SAP desde fila actual: {fila_actual}")
        
        # Obtener el número de filas antes de agregar
        filas_antes = grid.RowCount
        logger.info(f"📊 Filas antes de agregar: {filas_antes}")
        
        # Ejecutar el script exacto para agregar fila
        logger.info(f"🔧 Ejecutando script de agregar fila...")
        try:
            time.sleep(1)
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell( fila_actual,"")
            logger.info(f"✅ currentCellColumn ejecutado")
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = str(fila_actual)
            logger.info(f"✅ selectedRows ejecutado")
            # Verificar si el botón existe antes de presionarlo
            try:
                time.sleep(1)
                btn = session.findById("wnd[0]/tbar[1]/btn[7]")
                logger.info(f"🔍 Botón encontrado: {btn.text if hasattr(btn, 'text') else 'Sin texto'}")
                time.sleep(1)
                btn.press()
                time.sleep(1)
                logger.info(f"✅ btn[7] presionado")
            except Exception as e:
                logger.error(f"❌ Error con el botón btn[7]: {e}")
                # Intentar listar todos los botones disponibles
                try:
                    logger.info(f"🔍 Listando botones disponibles en tbar[1]:")
                    for i in range(10):
                        try:
                            btn_test = session.findById(f"wnd[0]/tbar[1]/btn[{i}]")
                            logger.info(f"   - btn[{i}]: {btn_test.text if hasattr(btn_test, 'text') else 'Sin texto'}")
                        except:
                            continue
                except Exception as e2:
                    logger.error(f"❌ Error listando botones: {e2}")
                return False, None
        except Exception as e:
            logger.error(f"❌ Error ejecutando script de agregar fila: {e}")
            return False, None
        
        # Esperar a que se agregue la fila y SAP actualice el grid
        logger.info(f"⏳ Esperando 3 segundos para que SAP actualice el grid...")
        time.sleep(3)
        
        # Obtener el número de filas después de agregar
        filas_despues = grid.RowCount
        logger.info(f"📊 Filas después de agregar: {filas_despues}")
        
        # Verificar si el grid se actualizó
        if filas_despues == filas_antes:
            logger.warning(f"⚠️ Grid no se actualizó. Intentando verificar estado...")
            # Intentar verificar si hay alguna fila nueva
            try:
                for i in range(filas_antes, filas_antes + 2):
                    test_value = grid.getCellValue(i, "ZZEAN13")
                    logger.info(f"🔍 Fila {i}: EAN='{test_value}'")
            except Exception as e:
                logger.info(f"🔍 No se puede acceder a fila {i}: {e}")
        
        if filas_despues > filas_antes:
            # La nueva fila siempre será la siguiente a la fila actual
            nueva_fila_index = fila_actual + 1
            logger.info(f"✅ Nueva fila agregada exitosamente. Índice de nueva fila: {nueva_fila_index}")
            return True, nueva_fila_index
        else:
            logger.error(f"❌ No se detectó incremento en el número de filas")
            return False, None
        
    except Exception as e:
        logger.error(f"❌ Error agregando fila en SAP: {e}")
        return False, None


def procesar_ean_repetido(grid, session, ean, filas_excel, cantidades, lotes, fechas_vencimiento):
    """
    Procesa un EAN que aparece en múltiples filas del Excel.
    
    Args:
        grid: Grid de SAP
        session: Sesión de SAP
        ean: EAN a procesar
        filas_excel: Lista de índices de filas del Excel
        cantidades: Lista de cantidades confirmadas
        lotes: Lista de lotes de estuche
        fechas_vencimiento: Lista de fechas de vencimiento
        
    Returns:
        bool: True si se procesó exitosamente, False en caso contrario
    """
    try:
        logger.info(f"🔄 Procesando EAN repetido: {ean}")
        logger.info(f"   - Filas Excel: {filas_excel}")
        logger.info(f"   - Cantidades: {cantidades}")
        logger.info(f"   - Lotes: {lotes}")
        
        # Validar cantidades antes de procesar
        total_cantidad_excel = sum(cantidades)
        es_valido, cantidad_sap, mensaje_validacion = validar_cantidades_ean_repetido(grid, ean, total_cantidad_excel)
        
        if not es_valido:
            logger.error(f"❌ Validación de cantidades falló para EAN {ean}: {mensaje_validacion}")
            return False
        
        logger.info(f"✅ Validación de cantidades exitosa: {mensaje_validacion}")
        
        # Buscar la fila original en SAP para este EAN
        fila_sap_original, cantidad_sap, mensaje = find_best_sap_row_for_ean(
            grid, ean, cantidades[0]  # Usar la primera cantidad como referencia
        )
        
        if fila_sap_original is None:
            logger.error(f"❌ No se encontró fila SAP para EAN {ean}")
            return False
        
        logger.info(f"📝 Fila SAP original encontrada: {fila_sap_original}")
        
        # Procesar la primera fila del Excel en la fila original de SAP
        try:
            logger.info(f"📝 Cargando primera fila del Excel en fila SAP {fila_sap_original}")
            grid.modifyCell(fila_sap_original, "CANTIDAD", str(int(cantidades[0])))
            time.sleep(0.5)
            grid.modifyCell(fila_sap_original, "CHARG", str(lotes[0]))
            time.sleep(0.5)
            fecha_venc = pd.to_datetime(fechas_vencimiento[0]).strftime("%d.%m.%Y")
            grid.modifyCell(fila_sap_original, "VENCIMIENTO", fecha_venc)
            time.sleep(0.5)
            logger.info(f"✅ Primera fila procesada exitosamente")
            
        except Exception as e:
            logger.error(f"❌ Error procesando primera fila: {e}")
            return False
        
        # Procesar las filas adicionales del Excel
        for i in range(1, len(filas_excel)):
            try:
                logger.info(f"📝 Procesando fila adicional {i+1} del Excel")
                
                # Agregar nueva fila en SAP
                logger.info(f"🔍 Estado del grid antes de agregar fila para EAN {ean}: {grid.RowCount} filas")
                exito, nueva_fila_sap = agregar_fila_sap(grid, session, fila_sap_original)
                if not exito:
                    logger.error(f"❌ No se pudo agregar fila adicional para EAN {ean}")
                    return False
                
                logger.info(f"📝 Nueva fila SAP creada: {nueva_fila_sap}")
                logger.info(f"🔍 Estado del grid después de agregar fila para EAN {ean}: {grid.RowCount} filas")
                
                # Verificar la cantidad que viene por defecto en la nueva fila
                try:
                    cantidad_por_defecto = grid.getCellValue(nueva_fila_sap, "CANT_PEND")
                    logger.info(f"📊 Nueva fila {nueva_fila_sap}: cantidad por defecto = {cantidad_por_defecto}")
                except Exception as e:
                    logger.warning(f"⚠️ No se pudo leer cantidad por defecto: {e}")
                
                # Cargar datos en la nueva fila
                grid.modifyCell(nueva_fila_sap, "CANTIDAD", str(int(cantidades[i])))
                time.sleep(0.5)
                grid.modifyCell(nueva_fila_sap, "CHARG", str(lotes[i]))
                time.sleep(0.5)
                fecha_venc = pd.to_datetime(fechas_vencimiento[i]).strftime("%d.%m.%Y")
                grid.modifyCell(nueva_fila_sap, "VENCIMIENTO", fecha_venc)
                time.sleep(0.5)
                
                logger.info(f"✅ Fila adicional {i+1} procesada exitosamente")
                
            except Exception as e:
                logger.error(f"❌ Error procesando fila adicional {i+1}: {e}")
                return False
        
        logger.info(f"✅ EAN repetido {ean} procesado completamente")
        return True
        
    except Exception as e:
        logger.error(f"❌ Error procesando EAN repetido {ean}: {e}")
        return False


def validar_eans_excel_en_sap(grid, df_excel, oc):
    """
    Valida que todos los EANs del Excel existan en el grid de SAP antes del procesamiento.

    Args:
        grid: Grid de SAP
        df_excel: DataFrame del Excel
        oc: Número de orden de compra
        
    Returns:
        tuple: (todos_encontrados, eans_faltantes, mensaje)
    """
    try:
        total_rows_sap = grid.RowCount
        eans_excel = set()
        eans_sap = set()
        eans_faltantes = []

        # Obtener todos los EANs del Excel
        logger.info(f"🔍 Leyendo EANs del Excel...")
        for _, row in df_excel.iterrows():
            ean_excel = str(row.get('EAN', '')).strip()
            if ean_excel:
                eans_excel.add(ean_excel)
                logger.info(f"   - Excel: EAN='{ean_excel}'")

        # Obtener todos los EANs de SAP
        for sap_idx in range(total_rows_sap):
            try:
                ean_sap = grid.getCellValue(sap_idx, "ZZEAN13").strip()
                if ean_sap:
                    eans_sap.add(ean_sap)
            except Exception as e:
                logger.warning(f"Error accediendo a fila SAP {sap_idx}: {e}")
                continue
        
        # Verificar qué EANs del Excel no están en SAP
        logger.info(f"🔍 Comparando EANs del Excel con SAP...")
        for ean_excel in eans_excel:
            logger.info(f"   - Verificando EAN Excel '{ean_excel}' en SAP...")
            if ean_excel in eans_sap:
                logger.info(f"     ✅ EAN '{ean_excel}' encontrado en SAP")
            else:
                logger.info(f"     ❌ EAN '{ean_excel}' NO encontrado en SAP")
                eans_faltantes.append(ean_excel)
        
        todos_encontrados = len(eans_faltantes) == 0
        
        if todos_encontrados:
            mensaje = f"✅ Todos los EANs del Excel ({len(eans_excel)}) encontrados en SAP"
        else:
            mensaje = f"❌ {len(eans_faltantes)} EANs del Excel no encontrados en SAP: {eans_faltantes}"
        
        logger.info(f"📊 Validación EAN para OC {oc}:")
        logger.info(f"   - EANs en Excel: {len(eans_excel)}")
        logger.info(f"   - EANs en SAP: {len(eans_sap)}")
        logger.info(f"   - EANs faltantes: {len(eans_faltantes)}")
        logger.info(f"   - Resultado: {mensaje}")
        
        return todos_encontrados, eans_faltantes, mensaje
        
    except Exception as e:
        logger.error(f"Error en validar_eans_excel_en_sap: {e}")
        return False, [], f"Error en validación: {e}"


def buscar_ean_en_sap_desde_fila(grid, ean_buscar, fila_inicio):
    """
    Busca un EAN en SAP desde una fila específica hacia adelante.
    
    Args:
        grid: Grid de SAP
        ean_buscar: EAN a buscar
        fila_inicio: Fila desde donde empezar a buscar
        
    Returns:
        int: Índice de la fila donde se encontró el EAN, o None si no se encontró
    """
    try:
        total_filas = grid.RowCount
        logger.info(f"🔍 Buscando EAN '{ean_buscar}' desde fila {fila_inicio} hasta {total_filas-1}")
        
        for fila in range(fila_inicio, total_filas):
            try:
                ean_sap = grid.getCellValue(fila, "ZZEAN13").strip()
                if ean_sap == ean_buscar.strip():
                    logger.info(f"✅ EAN '{ean_buscar}' encontrado en fila {fila}")
                    return fila
            except Exception as e:
                logger.warning(f"⚠️ Error accediendo a fila {fila}: {e}")
                continue
        
        logger.warning(f"❌ EAN '{ean_buscar}' no encontrado desde fila {fila_inicio}")
        return None
        
    except Exception as e:
        logger.error(f"❌ Error buscando EAN '{ean_buscar}': {e}")
        return None


def procesar_ean_secuencial_simple(grid, session, ean, filas_excel, cantidades, lotes, fechas_vencimiento):
    """
    Procesa un EAN de forma secuencial usando búsqueda desde fila 0.
    
    Args:
        grid: Grid de SAP
        session: Sesión de SAP
        ean: EAN a procesar
        filas_excel: Lista de índices de filas del Excel
        cantidades: Lista de cantidades confirmadas
        lotes: Lista de lotes de estuche
        fechas_vencimiento: Lista de fechas de vencimiento
        
    Returns:
        bool: True si se procesó exitosamente, False en caso contrario
    """
    try:
        logger.info(f"🔄 Procesando EAN secuencial simple: {ean}")
        logger.info(f"   - Filas Excel: {filas_excel}")
        logger.info(f"   - Cantidades: {cantidades}")
        logger.info(f"   - Lotes: {lotes}")
        
        # Si solo hay una fila, procesar normalmente
        if len(filas_excel) == 1:
            logger.info(f"📝 EAN {ean} tiene solo una fila, procesando normalmente")
            return True
        
        # Si hay múltiples filas, validar cantidades primero
        total_cantidad_excel = sum(cantidades)
        es_valido, cantidad_sap, mensaje_validacion = validar_cantidades_ean_repetido(grid, ean, total_cantidad_excel)
        
        if not es_valido:
            logger.error(f"❌ Validación de cantidades falló para EAN {ean}: {mensaje_validacion}")
            return False
        
        logger.info(f"✅ Validación de cantidades exitosa: {mensaje_validacion}")
        
        # Buscar la primera fila de SAP para este EAN
        fila_sap_actual = buscar_ean_en_sap_desde_fila(grid, ean, 0)
        
        if fila_sap_actual is None:
            logger.error(f"❌ No se encontró fila SAP para EAN {ean}")
            return False
        
        logger.info(f"📝 Primera fila SAP encontrada: {fila_sap_actual}")
        
        # Procesar la primera fila del Excel en la fila original de SAP
        try:
            logger.info(f"📝 Cargando primera fila del Excel en fila SAP {fila_sap_actual}")
            grid.modifyCell(fila_sap_actual, "CANTIDAD", str(int(cantidades[0])))
            time.sleep(0.5)
            grid.modifyCell(fila_sap_actual, "CHARG", str(lotes[0]))
            time.sleep(0.5)
            fecha_venc = pd.to_datetime(fechas_vencimiento[0]).strftime("%d.%m.%Y")
            grid.modifyCell(fila_sap_actual, "VENCIMIENTO", fecha_venc)
            time.sleep(0.5)
            logger.info(f"✅ Primera fila procesada exitosamente")
            
        except Exception as e:
            logger.error(f"❌ Error procesando primera fila: {e}")
            return False
        
        # Procesar las filas adicionales del Excel (lotes adicionales)
        for i in range(1, len(filas_excel)):
            try:
                logger.info(f"📝 Procesando lote adicional {i+1} del Excel para EAN {ean}")
                
                # Agregar nueva fila en SAP
                logger.info(f"🔍 Estado del grid antes de agregar fila para EAN {ean}: {grid.RowCount} filas")
                exito, nueva_fila_sap = agregar_fila_sap(grid, session, fila_sap_actual)
                if not exito:
                    logger.error(f"❌ No se pudo agregar fila adicional para EAN {ean}")
                    return False
                
                logger.info(f"📝 Nueva fila SAP creada: {nueva_fila_sap}")
                logger.info(f"🔍 Estado del grid después de agregar fila para EAN {ean}: {grid.RowCount} filas")
                
                # Verificar la cantidad que viene por defecto en la nueva fila
                try:
                    cantidad_por_defecto = grid.getCellValue(nueva_fila_sap, "CANT_PEND")
                    logger.info(f"📊 Nueva fila {nueva_fila_sap}: cantidad por defecto = {cantidad_por_defecto}")
                except Exception as e:
                    logger.warning(f"⚠️ No se pudo leer cantidad por defecto: {e}")
                
                # Cargar datos en la nueva fila
                grid.modifyCell(nueva_fila_sap, "CANTIDAD", str(int(cantidades[i])))
                time.sleep(0.5)
                grid.modifyCell(nueva_fila_sap, "CHARG", str(lotes[i]))
                time.sleep(0.5)
                fecha_venc = pd.to_datetime(fechas_vencimiento[i]).strftime("%d.%m.%Y")
                grid.modifyCell(nueva_fila_sap, "VENCIMIENTO", fecha_venc)
                time.sleep(0.5)
                
                logger.info(f"✅ Lote adicional {i+1} procesado exitosamente")
                
            except Exception as e:
                logger.error(f"❌ Error procesando lote adicional {i+1}: {e}")
                return False
        
        logger.info(f"✅ EAN {ean} procesado completamente con {len(filas_excel)} lotes")
        return True
        
    except Exception as e:
        logger.error(f"❌ Error procesando EAN secuencial simple {ean}: {e}")
        return False


def process_entrega(session, path_excel, oc):
    """
    Procesa un Excel y carga dinámicamente los datos en SAP GUI.
    Implementa validación exhaustiva de EAN: busca cada EAN del Excel en todas las filas de SAP
    y carga los datos en la fila correcta una vez identificada.
    
    MANEJO DE ERRORES ROBUSTO:
    - Si cualquier error ocurre durante el procesamiento, el archivo se mueve a carpeta de errores
    - Se registra el error en un archivo de log específico
    - Se continúa con la siguiente orden de compra
    """
    import pandas as pd
    import time
    from utils import consultarCadenaFrio
    import logging
    import traceback
    logger = logging.getLogger(__name__)

    try:
        logger.info(f"🚀 Iniciando procesamiento de OC {oc} - Archivo: {path_excel}")
        
        # 1. Leer Excel y filtrar filas válidas
        df = pd.read_excel(path_excel)
        df = df[df['Fecha Vencimiento'].notna()]
        if df.empty:
            error_msg = f"El archivo {path_excel} no contiene filas válidas."
            logger.error(error_msg)
            if os.path.exists(path_excel):
                exito = mover_archivo_a_errores(path_excel, oc, error_msg)
                if exito:
                    logger.info(f"✅ Archivo movido exitosamente a errores")
                else:
                    logger.error(f"❌ Error moviendo archivo a errores")
            else:
                logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
            return

        # 2. Extraer remito
        row = df.iloc[0]
        remito_completo = row['Remito y Nro. Entrega']
        
        # Separar el remito del número de entrega
        partes = remito_completo.split()
        remito_con_r = partes[0]  # "0114R02179687" o "R0114"
        
        # Extraer las dos partes del remito
        if "R" in remito_con_r:
            if remito_con_r.startswith("R"):
                # Caso: "R0114" - la R está al principio
                remito1 = remito_con_r[1:5]  # "0114"
                remito2 = remito_con_r[5:] if len(remito_con_r) > 5 else ""
            else:
                # Caso: "0114R02179687" - la R está en el medio
                partes_remito = remito_con_r.split("R")
                remito1 = partes_remito[0]  # "0114"
                remito2 = partes_remito[1]  # "02179687"
        else:
            # Fallback si no hay "R" en el formato
            logger.warning(f"Formato de remito inesperado: {remito_con_r}")
            remito1 = remito_con_r[:4]
            remito2 = remito_con_r[4:] if len(remito_con_r) > 4 else ""
        
        # Log para verificar la extracción correcta
        logger.info(f"Remito completo: {remito_completo}")
        logger.info(f"Remito1 extraído: {remito1}")
        logger.info(f"Remito2 extraído: {remito2}")
        
        # Validar que los datos extraídos son válidos
        if not remito1 or not remito2:
            error_msg = f"Error en la extracción del remito. Remito1: '{remito1}', Remito2: '{remito2}'"
            logger.error(error_msg)
            logger.error(f"Remito completo: {remito_completo}")
            if os.path.exists(path_excel):
                exito = mover_archivo_a_errores(path_excel, oc, error_msg)
                if exito:
                    logger.info(f"✅ Archivo movido exitosamente a errores")
                else:
                    logger.error(f"❌ Error moviendo archivo a errores")
            else:
                logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
            return

        # 3. Navegar a la transacción SAP
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nZMM_RECEP_DOCU"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Cargar OC en SAP
        session.findById("wnd[0]/usr/ctxtSO_EBELN-LOW").text = oc
        session.findById("wnd[0]/usr/ctxtSO_EBELN-LOW").caretPosition = len(oc)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(1)
        session.findById("wnd[0]/tbar[1]/btn[20]").press()
        time.sleep(1)

        # 4. Consultar cadena de frio
        try:
            frio = consultarCadenaFrio(oc)
        except Exception as e:
            logger.warning(f"No se pudo consultar cadena de frio: {e}")
            frio = False

        # 5. VALIDACIÓN EXHAUSTIVA DE EAN Y CARGA DE DATOS
        grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        total_rows_sap = grid.RowCount
        logger.info(f"📊 Grid SAP tiene {total_rows_sap} filas")
        
        # VALIDACIÓN PREVIA: Verificar que todos los EANs del Excel existan en SAP
        logger.info(f"🔍 Iniciando validación previa de EANs para OC {oc}")
        todos_encontrados, eans_faltantes, mensaje_validacion = validar_eans_excel_en_sap(grid, df, oc)
        
        if not todos_encontrados:
            logger.error(f"❌ {mensaje_validacion}")
            # Registrar error para cada EAN faltante
            for ean_faltante in eans_faltantes:
                registrar_error_ean_no_encontrado(oc, ean_faltante, path_excel)
            logger.error(f"❌ Abortando procesamiento de OC {oc} debido a EANs faltantes")
            
            # Mover archivo a errores por EANs faltantes
            error_msg = f"EANs faltantes en SAP: {eans_faltantes}"
            if os.path.exists(path_excel):
                exito = mover_archivo_a_errores(path_excel, oc, error_msg)
                if exito:
                    logger.info(f"✅ Archivo movido exitosamente a errores por EANs faltantes")
                else:
                    logger.error(f"❌ Error moviendo archivo a errores")
            else:
                logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
            return
        
        logger.info(f"✅ {mensaje_validacion}")
        
        # Contadores para logging
        eans_encontrados = 0
        eans_no_encontrados = 0
        filas_procesadas = 0
        eans_con_error = []
        
        # Procesar cada EAN del Excel secuencialmente
        eans_procesados = set()  # Para evitar procesar el mismo EAN múltiples veces
        
        for excel_idx, excel_row in df.iterrows():
            ean_excel = str(excel_row.get('EAN', '')).strip()
            logger.info(f"🔍 Procesando fila Excel {excel_idx}: EAN='{ean_excel}'")
            
            # Evitar procesar el mismo EAN múltiples veces
            if ean_excel in eans_procesados:
                logger.info(f"⏭️ EAN {ean_excel} ya procesado, saltando...")
                continue
            
            # Verificar si este EAN tiene filas adicionales (es repetido)
            filas_mismo_ean = df[df['EAN'].astype(str).str.strip() == ean_excel.strip()]
            
            logger.info(f"🔍 EAN '{ean_excel}' encontrado en {len(filas_mismo_ean)} filas del Excel")
            
            if len(filas_mismo_ean) > 1:
                # EAN repetido en Excel - procesar múltiples lotes
                logger.info(f"🔄 EAN repetido detectado: {ean_excel} con {len(filas_mismo_ean)} lotes en Excel")
                
                # Preparar datos para procesamiento secuencial
                filas_excel = filas_mismo_ean.index.tolist()
                cantidades = filas_mismo_ean['Cant confirmada'].tolist()
                lotes = filas_mismo_ean['Lote estuche'].tolist()
                fechas_vencimiento = filas_mismo_ean['Fecha Vencimiento'].tolist()
                
                # Procesar EAN repetido usando búsqueda secuencial
                if procesar_ean_secuencial_simple(grid, session, ean_excel, filas_excel, cantidades, lotes, fechas_vencimiento):
                    logger.info(f"✅ EAN repetido {ean_excel} procesado con éxito.")
                    eans_encontrados += 1
                    filas_procesadas += len(filas_mismo_ean)
                    eans_procesados.add(ean_excel)
                else:
                    logger.error(f"❌ Error procesando EAN repetido {ean_excel}.")
                    eans_con_error.append(ean_excel)
                    registrar_error_ean_repetido(oc, ean_excel, "Error en procesamiento de EAN repetido", path_excel)
                
            else:
                # EAN individual en Excel - procesar normalmente
                logger.info(f"📝 EAN individual detectado: {ean_excel}")
                cantidad_confirmada = str(int(excel_row['Cant confirmada']))
                lote_estuche = excel_row['Lote estuche']
                fecha_vencimiento = pd.to_datetime(excel_row['Fecha Vencimiento'], dayfirst=True).strftime("%d.%m.%Y")
                
                # Buscar este EAN en SAP desde fila 0
                fila_sap_encontrada = buscar_ean_en_sap_desde_fila(grid, ean_excel, 0)
                
                if fila_sap_encontrada is None:
                    logger.error(f"❌ EAN '{ean_excel}' no encontrado en SAP")
                    eans_no_encontrados += 1
                    eans_con_error.append(ean_excel)
                    continue
                
                # Cargar datos en la fila encontrada
                try:
                    logger.info(f"📝 Cargando datos en fila SAP {fila_sap_encontrada}")
                    grid.modifyCell(fila_sap_encontrada, "CANTIDAD", cantidad_confirmada)
                    time.sleep(0.5)
                    grid.modifyCell(fila_sap_encontrada, "CHARG", lote_estuche)
                    time.sleep(0.5)
                    grid.modifyCell(fila_sap_encontrada, "VENCIMIENTO", fecha_vencimiento)
                    time.sleep(0.5)
                    filas_procesadas += 1
                    eans_encontrados += 1
                    logger.info(f"✅ Datos cargados exitosamente en fila SAP {fila_sap_encontrada}")
                    eans_procesados.add(ean_excel)
                    
                except Exception as e:
                    logger.error(f"❌ Error cargando datos en fila SAP {fila_sap_encontrada}: {e}")
                    continue
        
        logger.info(f"🔍 Bucle de procesamiento completado. Total filas Excel: {len(df)}")
        logger.info(f"🔍 EANs procesados: {eans_procesados}")
        
        # Resumen del procesamiento
        logger.info(f"📊 RESUMEN PROCESAMIENTO:")
        logger.info(f"   - EANs encontrados: {eans_encontrados}")
        logger.info(f"   - EANs no encontrados: {eans_no_encontrados}")
        logger.info(f"   - Filas procesadas: {filas_procesadas}")
        
        # Si hay EANs no encontrados, registrar error y abortar
        if eans_con_error:
            logger.error(f"❌ OC {oc} tiene EANs no encontrados: {eans_con_error}")
            for ean_error in eans_con_error:
                registrar_error_ean_no_encontrado(oc, ean_error, path_excel)
            logger.error(f"❌ Abortando procesamiento de OC {oc} debido a EANs no encontrados")
            
            # Mover archivo a errores por EANs no encontrados
            error_msg = f"EANs no encontrados en SAP: {eans_con_error}"
            if os.path.exists(path_excel):
                exito = mover_archivo_a_errores(path_excel, oc, error_msg)
                if exito:
                    logger.info(f"✅ Archivo movido exitosamente a errores por EANs no encontrados")
                else:
                    logger.error(f"❌ Error moviendo archivo a errores")
            else:
                logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
            return
        
        # Si no se procesó ninguna fila, abortar
        if filas_procesadas == 0:
            logger.error(f"❌ No se pudo procesar ninguna fila del Excel. Abortando.")
            
            # Mover archivo a errores por falta de procesamiento
            error_msg = "No se pudo procesar ninguna fila del Excel"
            if os.path.exists(path_excel):
                exito = mover_archivo_a_errores(path_excel, oc, error_msg)
                if exito:
                    logger.info(f"✅ Archivo movido exitosamente a errores por falta de procesamiento")
                else:
                    logger.error(f"❌ Error moviendo archivo a errores")
            else:
                logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
            return
        
        # Presionar Enter para confirmar cambios
        grid.pressEnter()
        time.sleep(1)
        
    except Exception as e:
        logger.critical(f"Error al cargar datos en la grilla SAP: {e}")
        
        # Mover archivo a errores por error en carga de datos
        error_msg = f"Error al cargar datos en la grilla SAP: {str(e)}"
        if os.path.exists(path_excel):
            exito = mover_archivo_a_errores(path_excel, oc, error_msg)
            if exito:
                logger.info(f"✅ Archivo movido exitosamente a errores por error en carga de datos")
            else:
                logger.error(f"❌ Error moviendo archivo a errores")
        else:
            logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
        return

    # 6. Completar datos de remito y bultos
    try:
        session.findById("wnd[0]/tbar[1]/btn[21]").press()
        session.findById("wnd[1]/usr/txtGV_0100_REMITO1").text = remito1
        time.sleep(0.5)
        session.findById("wnd[1]/usr/txtGV_0100_REMITO2").text = remito2
        time.sleep(0.5)
        
        try:
            session.findById("wnd[1]/usr/txtGV_0100_BULTOS_FRIO").text = "1"
            time.sleep(0.5)
        except Exception:
            logger.info("No hay campo BULTOS_SECO")
        try:
            session.findById("wnd[1]/usr/txtGV_0100_BULTOS_SECO").text = "1"
            time.sleep(0.5)
        except Exception:
            logger.info("No hay campo BULTOS_FRIO")

        session.findById("wnd[1]/usr/txtGV_0100_FACTURA1").setFocus()
        session.findById("wnd[1]/usr/txtGV_0100_FACTURA1").caretPosition = 0
        time.sleep(0.5)
        session.findById("wnd[1]/usr/btnBOT_GENERAR").press()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)
        remito = f"R{remito1+remito2}"
        session.findById("wnd[1]/usr/txtSSFPP-TDCOVTITLE").text = remito
        session.findById("wnd[1]/tbar[0]/btn[86]").press()
        logger.info(f"✅ Entrega creada en SAP para OC {oc}")
        
        # Esperar un poco para que se genere el PDF completamente
        time.sleep(3)
        
        # Renombrar el archivo PDF de la etiqueta
        logger.info(f"🔄 Iniciando renombrado de PDF para OC {oc}")
        # Buscar en la carpeta de Downloads donde SAP guarda los PDFs
        carpeta_pdfs = r"C:\Users\recepcion1\Documents\Etiquetas Entregas Entrantes Farmanet"
        renombrar_pdf_etiqueta(remito, carpeta_pdfs)
        
        logger.info(f"✅ Procesamiento completado exitosamente para OC {oc}")
        
    except Exception as e:
        # MANEJO DE ERRORES GLOBAL - Cualquier error no capturado
        error_msg = f"Error crítico en procesamiento de SAP para OC {oc}: {str(e)}"
        logger.critical(error_msg)
        logger.critical(f"Traceback completo: {traceback.format_exc()}")
        
        # Verificar que el archivo existe antes de moverlo
        if os.path.exists(path_excel):
            logger.info(f"📁 Moviendo archivo a errores: {path_excel}")
            # Mover archivo a carpeta de errores
            exito = mover_archivo_a_errores(path_excel, oc, error_msg)
            if exito:
                logger.info(f"✅ Archivo movido exitosamente a errores")
            else:
                logger.error(f"❌ Error moviendo archivo a errores")
        else:
            logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
        
        logger.info(f"🔄 Continuando con la siguiente orden de compra...")
        return


def crear_resumen_eans_repetidos(oc, eans_repetidos_excel, path_excel):
    """
    Crea un archivo de resumen detallado del procesamiento de EANs repetidos.
    
    Args:
        oc: Número de orden de compra
        eans_repetidos_excel: Diccionario con información de EANs repetidos
        path_excel: Ruta del archivo Excel
    """
    try:
        import os
        from datetime import datetime
        
        # Crear directorio de resúmenes si no existe
        resumen_dir = os.path.join(os.getcwd(), "Resumenes")
        os.makedirs(resumen_dir, exist_ok=True)
        
        # Crear archivo de resumen
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        resumen_file = os.path.join(resumen_dir, f"resumen_eans_repetidos_{oc}_{timestamp}.txt")
        
        with open(resumen_file, "w", encoding="utf-8") as f:
            f.write("RESUMEN DE EANs REPETIDOS PROCESADOS\n")
            f.write("=" * 50 + "\n")
            f.write(f"OC: {oc}\n")
            f.write(f"Fecha: {datetime.now()}\n")
            f.write(f"Archivo Excel: {path_excel}\n")
            f.write(f"Total EANs repetidos: {len(eans_repetidos_excel)}\n")
            f.write("=" * 50 + "\n\n")
            
            for ean, info in eans_repetidos_excel.items():
                f.write(f"EAN: {ean}\n")
                f.write(f"  - Filas Excel: {info['filas']}\n")
                f.write(f"  - Cantidades: {info['cantidades']}\n")
                f.write(f"  - Lotes: {info['lotes']}\n")
                f.write(f"  - Total cantidad: {info['total_cantidad']}\n")
                f.write(f"  - Fechas vencimiento: {[pd.to_datetime(fecha).strftime('%d.%m.%Y') for fecha in info['fechas_vencimiento']]}\n")
                f.write("\n")
            
            f.write("=" * 50 + "\n")
            f.write("PROCESAMIENTO COMPLETADO EXITOSAMENTE\n")
            f.write("Todos los EANs repetidos fueron procesados correctamente.\n")
        
        logger.info(f"📋 Resumen de EANs repetidos creado en: {resumen_file}")
        
    except Exception as e:
        logger.error(f"❌ Error creando resumen de EANs repetidos para OC {oc}: {e}")


def mover_archivo_a_errores(path_excel, oc, error_descripcion):
    """
    Mueve un archivo Excel a la carpeta de errores para evitar reprocesamiento.
    Una OC puede tener múltiples entregas, por lo que movemos específicamente
    el archivo de la entrega que falló.
    
    Args:
        path_excel: Ruta del archivo Excel
        oc: Número de orden de compra
        error_descripcion: Descripción del error que causó el fallo
    """
    try:
        # Crear directorio de errores si no existe
        error_dir = os.path.join(os.getcwd(), "Errores")
        os.makedirs(error_dir, exist_ok=True)
        
        # Crear subdirectorio para archivos no procesados
        no_procesados_dir = os.path.join(error_dir, "No_Procesados")
        os.makedirs(no_procesados_dir, exist_ok=True)
        
        # Extraer información del archivo para identificar la entrega específica
        nombre_archivo = os.path.basename(path_excel)
        nombre_sin_extension = os.path.splitext(nombre_archivo)[0]
        
        # Generar nombre que preserve la identificación de la entrega
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nuevo_nombre = f"{nombre_sin_extension}_ERROR_{timestamp}.xlsx"
        
        # Ruta destino
        ruta_destino = os.path.join(no_procesados_dir, nuevo_nombre)
        
        # Verificar si el archivo existe antes de moverlo
        if not os.path.exists(path_excel):
            logger.warning(f"⚠️ Archivo no encontrado para mover a errores: {path_excel}")
            return False
        
        logger.info(f"🔍 Archivo encontrado, procediendo a mover: {path_excel}")
        logger.info(f"📂 Destino: {ruta_destino}")
        
        # Mover el archivo
        try:
            shutil.move(path_excel, ruta_destino)
            logger.info(f"✅ Archivo movido exitosamente")
            
            # Verificar que el archivo se movió correctamente
            if os.path.exists(ruta_destino):
                logger.info(f"✅ Verificación: archivo existe en destino")
            else:
                logger.error(f"❌ Error: archivo no existe en destino después del movimiento")
                return False
                
        except Exception as e:
            logger.error(f"❌ Error moviendo archivo: {e}")
            return False
        
        # Crear archivo de log del error con información específica de la entrega
        log_error_file = os.path.join(no_procesados_dir, f"error_procesamiento_{oc}_{timestamp}.txt")
        
        with open(log_error_file, "w", encoding="utf-8") as f:
            f.write("ERROR EN PROCESAMIENTO DE SAP - ENTREGA NO PROCESADA\n")
            f.write("=" * 60 + "\n")
            f.write(f"OC: {oc}\n")
            f.write(f"Archivo Original: {nombre_archivo}\n")
            f.write(f"Fecha Error: {datetime.now()}\n")
            f.write(f"Archivo Movido a: {ruta_destino}\n")
            f.write(f"Error: {error_descripcion}\n")
            f.write("=" * 60 + "\n")
            f.write("\nESTA ENTREGA ESPECÍFICA HA SIDO MOVIDA A ERRORES.\n")
            f.write("No se volverá a procesar automáticamente.\n")
            f.write("Revisar manualmente antes de reprocesar.\n")
            f.write("Se continúa con la siguiente entrega/OC.\n")
        
        logger.info(f"📁 Archivo movido a errores: {ruta_destino}")
        logger.info(f"📝 Log de error creado: {log_error_file}")
        logger.info(f"🔄 Esta entrega específica no se reprocesará automáticamente")
        
        return True
        
    except Exception as e:
        logger.error(f"❌ Error moviendo archivo a carpeta de errores: {e}")
        return False


def verificar_archivo_en_errores(nombre_archivo):
    """
    Verifica si un archivo ya está en la carpeta de errores para evitar reprocesamiento.
    
    Args:
        nombre_archivo: Nombre del archivo a verificar
        
    Returns:
        bool: True si el archivo está en errores, False en caso contrario
    """
    try:
        # Buscar en la carpeta de errores
        error_dir = os.path.join(os.getcwd(), "Errores", "No_Procesados")
        
        if not os.path.exists(error_dir):
            return False
        
        # Buscar archivos que contengan el nombre base del archivo
        nombre_base = os.path.splitext(nombre_archivo)[0]
        
        for archivo in os.listdir(error_dir):
            if archivo.startswith(nombre_base) and archivo.endswith('.xlsx'):
                logger.info(f"⚠️ Archivo ya está en errores: {archivo}")
                return True
        
        return False
        
    except Exception as e:
        logger.error(f"❌ Error verificando archivo en errores: {e}")
        return False


def cerrar_sap(sesionsap):
    SapGuiAuto = win32com.client.GetObject('SAPGUI')
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return
    application = SapGuiAuto.GetScriptingEngine

    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        return
    session = connection.Children(sesionsap)
    if not type(session) == win32com.client.CDispatch:
        return
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
        session.findById("wnd[0]").sendVKey(0)
        return True
    except Exception:
        pass


def renombrar_pdf_etiqueta(remito, carpeta):
    """
    Busca un archivo PDF que contenga el remito en la carpeta especificada y lo renombra.
    
    Args:
        remito: String del remito a buscar (ej: "R011402180514")
        carpeta: Ruta de la carpeta donde buscar
        
    Returns:
        bool: True si se renombró exitosamente, False en caso contrario
    """
    try:
        import os
        import glob
        import shutil
        from datetime import datetime
        
        logger.info(f"🔍 Buscando PDF con remito '{remito}' en carpeta: {carpeta}")
        
        # Verificar que la carpeta existe
        if not os.path.exists(carpeta):
            logger.warning(f"⚠️ Carpeta no encontrada: {carpeta}")
            return False
        
        # Buscar archivos PDF en la carpeta
        patron_pdf = os.path.join(carpeta, "*.pdf")
        archivos_pdf = glob.glob(patron_pdf)
        
        if not archivos_pdf:
            logger.warning(f"⚠️ No se encontraron archivos PDF en la carpeta: {carpeta}")
            return False
        
        # Buscar archivo que contenga el remito
        archivo_encontrado = None
        for archivo in archivos_pdf:
            nombre_archivo = os.path.basename(archivo)
            if remito in nombre_archivo:
                archivo_encontrado = archivo
                logger.info(f"✅ Archivo encontrado: {nombre_archivo}")
                break
        
        if not archivo_encontrado:
            logger.info(f"ℹ️ No se encontró archivo que contenga el remito '{remito}'")
            return False
        
        # Generar nuevo nombre con solo el remito
        extension = os.path.splitext(archivo_encontrado)[1]
        nuevo_nombre = f"{remito}{extension}"
        ruta_destino = os.path.join(carpeta, nuevo_nombre)
        
        # Verificar si el archivo de destino ya existe
        if os.path.exists(ruta_destino):
            logger.warning(f"⚠️ El archivo de destino ya existe: {nuevo_nombre}")
            return False
        
        try:
            # Renombrar el archivo
            os.rename(archivo_encontrado, ruta_destino)
            
            logger.info(f"✅ Archivo renombrado exitosamente:")
            logger.info(f"   - Original: {os.path.basename(archivo_encontrado)}")
            logger.info(f"   - Nuevo: {nuevo_nombre}")
            logger.info(f"   - Remito: {remito}")
            
            return True
            
        except Exception as e:
            logger.error(f"❌ Error renombrando archivo: {e}")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error en renombrar_pdf_etiqueta: {e}")
        return False
