#!/usr/bin/env python3
"""
Script de prueba para verificar la conexión con SAP
"""

import os
import sys
import logging

# Agregar rutas para imports
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'bot_farmanet'))
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'shared'))

from abrirsap import ingresarsap
from sap import get_sap_session

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - TEST_SAP - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def test_sap_connection():
    """Prueba la conexión con SAP"""
    logger.info("=== PRUEBA DE CONEXIÓN SAP ===")
    
    try:
        # 1. Abrir SAP
        logger.info("🔧 Abriendo SAP GUI...")
        ingresarsap("PRD", "cprosianiuk", "Scienza2025Scienza2025#")
        logger.info("✅ SAP abierto correctamente")
        
        # 2. Obtener sesión
        logger.info("🔑 Obteniendo sesión de SAP...")
        sap_session = get_sap_session()
        logger.info("✅ Sesión de SAP obtenida correctamente")
        
        # 3. Verificar que la sesión es válida
        if sap_session:
            logger.info("✅ Conexión con SAP exitosa")
            logger.info(f"📋 Información de sesión: {sap_session}")
            return True
        else:
            logger.error("❌ No se pudo obtener sesión de SAP")
            return False
            
    except Exception as e:
        logger.error(f"❌ Error en conexión con SAP: {e}")
        return False

if __name__ == "__main__":
    success = test_sap_connection()
    if success:
        print("\n✅ Prueba de conexión SAP exitosa")
    else:
        print("\n❌ Prueba de conexión SAP fallida") 