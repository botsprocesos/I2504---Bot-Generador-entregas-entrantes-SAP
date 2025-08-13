#!/usr/bin/env python3
"""
Bot SAP Processor - Procesa entregas entrantes desde Excel files
Sensa la carpeta no_procesados y genera entregas entrantes en SAP
"""

import os
import sys
import time
import logging
import schedule
from datetime import datetime
from pathlib import Path
import re

# Agregar el directorio padre al path para importar módulos
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

# Agregar rutas para imports
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'shared'))
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'bot_farmanet'))

# Importar módulos del bot
from sap import process_entrega, get_sap_session
from utils import setup_logging, ensure_directories

# Importar módulos de SAP
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'bot_farmanet'))
from abrirsap import ingresarsap

def extraer_numero_oc(filename):
    """
    Extrae el número de OC del inicio del nombre del archivo.
    
    Formato esperado: "NúmeroOC NúmeroEntrega"
    Ejemplos:
    - "5600025440 0082214777" -> "5600025440"
    - "5100064001 0082214479" -> "5100064001"
    - "5600025443 0082214783" -> "5600025443"
    
    Args:
        filename (str): Nombre del archivo sin extensión
        
    Returns:
        str: Número de OC extraído o None si no se puede extraer
    """
    try:
        # Buscar el primer número al inicio del archivo (antes del espacio)
        match = re.match(r'^(\d+)', filename)
        if match:
            oc_number = match.group(1)
            # Verificar que sea un número razonable (entre 8 y 12 dígitos)
            if 8 <= len(oc_number) <= 12:
                return oc_number
        
        # Si no encuentra al inicio, buscar el primer número en el archivo
        matches = re.findall(r'\b(\d{8,12})\b', filename)
        if matches:
            return matches[0]  # Retornar el primer número encontrado
        
        return None
        
    except Exception as e:
        logging.error(f"Error extrayendo OC de {filename}: {e}")
        return None

def setup_logging_sap():
    """Configurar logging específico para el bot SAP"""
    log_dir = Path(__file__).parent.parent / "Logs"
    log_dir.mkdir(exist_ok=True)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - SAP_PROCESSOR - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / f"sap_processor_{datetime.now().strftime('%Y%m%d')}.log"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def ensure_directories_sap():
    """Crear directorios necesarios para el bot SAP"""
    base_dir = Path(__file__).parent.parent
    
    # Directorios necesarios
    directories = [
        base_dir / "no_procesados",
        base_dir / "Errores" / "SAP_Processor",
        base_dir / "Logs"
    ]
    
    for directory in directories:
        directory.mkdir(parents=True, exist_ok=True)
        logger.info(f"Directorio creado/verificado: {directory}")

def procesar_excel_files():
    """Procesar todos los Excel files en la carpeta no_procesados"""
    logger.info("🔍 Iniciando procesamiento de Excel files...")
    
    # PRIMERO: Abrir SAP y autenticarse
    logger.info("🔧 Abriendo SAP GUI...")
    try:
        ingresarsap("PRD", "cprosianiuk", "Scienza2025Scienza2025#")
        sap_session = get_sap_session()
        logger.info("✅ SAP abierto y autenticado correctamente")
    except Exception as e:
        logger.error(f"❌ Error abriendo SAP: {e}")
        return
    
    base_dir = Path(__file__).parent.parent
    no_procesados_dir = base_dir / "no_procesados"
    errores_dir = base_dir / "Errores" / "SAP_Processor"
    
    # Verificar que existe la carpeta
    if not no_procesados_dir.exists():
        logger.warning("⚠️ Carpeta no_procesados no existe")
        return
    
    # Buscar archivos Excel
    excel_files = list(no_procesados_dir.glob("*.xlsx")) + list(no_procesados_dir.glob("*.xls"))
    
    if not excel_files:
        logger.info("📭 No hay archivos Excel para procesar")
        return
    
    logger.info(f"📁 Encontrados {len(excel_files)} archivos Excel para procesar")
    
    for excel_file in excel_files:
        try:
            logger.info(f"🔄 Procesando: {excel_file.name}")
            
            # Extraer OC del nombre del archivo (soporta múltiples formatos)
            filename = excel_file.stem
            oc_number = extraer_numero_oc(filename)
            
            if not oc_number:
                logger.warning(f"⚠️ No se pudo extraer OC de: {filename}")
                continue
            
            logger.info(f"📋 OC identificada: {oc_number}")
            
            # Procesar la entrega usando la sesión de SAP
            success = process_entrega(sap_session, str(excel_file), oc_number)
            
            if success:
                logger.info(f"✅ Entrega procesada exitosamente: {excel_file.name}")
                # Mover archivo procesado a carpeta de éxito o eliminarlo
                excel_file.unlink()
            else:
                logger.error(f"❌ Error procesando entrega: {excel_file.name}")
                # Mover a carpeta de errores
                error_path = errores_dir / excel_file.name
                excel_file.rename(error_path)
                logger.info(f"📁 Archivo movido a errores: {error_path}")
                
        except Exception as e:
            logger.error(f"❌ Error procesando {excel_file.name}: {str(e)}")
            # Mover a carpeta de errores
            try:
                error_path = errores_dir / excel_file.name
                excel_file.rename(error_path)
                logger.info(f"📁 Archivo movido a errores: {error_path}")
            except Exception as move_error:
                logger.error(f"❌ Error moviendo archivo a errores: {str(move_error)}")

def job_sap_processor():
    """Job principal del bot SAP Processor"""
    logger.info("🚀 Iniciando Bot SAP Processor")
    try:
        procesar_excel_files()
        logger.info("✅ Bot SAP Processor completado")
    except Exception as e:
        logger.error(f"❌ Error en Bot SAP Processor: {str(e)}")

def schedule_sap_processor():
    """Programar ejecución del bot SAP Processor"""
    # Ejecutar cada 5 minutos
    schedule.every(5).minutes.do(job_sap_processor)
    
    logger.info("⏰ Bot SAP Processor programado - ejecutándose cada 5 minutos")
    logger.info("🔄 Para detener: Ctrl+C")
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(30)  # Verificar cada 30 segundos
    except KeyboardInterrupt:
        logger.info("Bot SAP Processor detenido por el usuario")

if __name__ == "__main__":
    # Configurar logging
    logger = setup_logging_sap()
    
    # Crear directorios
    ensure_directories_sap()
    
    # Ejecutar automáticamente cada 5 minutos
    logger.info("Bot SAP Processor iniciado - ejecutándose automáticamente cada 5 minutos")
    schedule_sap_processor() 