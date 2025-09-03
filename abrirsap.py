
from typing import List
from schedule import logger
import os.path, sys, win32com.client,pythoncom, win32com.client, time, subprocess
import logging
import psutil

def kill_zombie_saplogon():
    """Mata procesos saplogon.exe que no tengan ruta o estén colgados."""
    logger = logging.getLogger(__name__)
    killed = 0
    for proc in psutil.process_iter(['name', 'exe', 'pid']):
        try:
            if proc.info['name'] and 'saplogon.exe' in proc.info['name'].lower():
                # Si no tiene ruta o está colgado
                if not proc.info['exe'] or not os.path.exists(proc.info['exe']):
                    logger.warning(f"Matando proceso zombie saplogon.exe PID {proc.info['pid']}")
                    proc.kill()
                    killed += 1
        except Exception as e:
            logger.warning(f"Error matando proceso zombie: {e}")
    return killed

def ingresarsap(amb, u, c, max_retries=10, wait_between=2):
    """
    Abre SAP GUI si no está abierto, espera a que esté listo para scripting y realiza login.
    Mata procesos zombie antes de abrir. Retorna True si tuvo éxito, False si no.
    """
    logger = logging.getLogger(__name__)
    try:
        pythoncom.CoInitialize()
        path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe" #RISE SAP 800
        # 1. Matar procesos zombie
        # kill_zombie_saplogon()
        # 2. Intentar obtener objeto COM SAPGUI
        logger.info("Abriendo SAP...")
        subprocess.Popen(path)
        for i in range(max_retries):
            try:
                SapGuiAuto = win32com.client.GetObject('SAPGUI')
                if SapGuiAuto:
                    
                    logger.info("SAP GUI scripting disponible tras lanzar SAP GUI.")
                    break
            except Exception:
                logger.info(f"Esperando SAP GUI scripting tras lanzar, intento {i+1}/{max_retries}")
                time.sleep(wait_between)
        if not SapGuiAuto:
            logger.error("No se pudo obtener el objeto SAPGUI para scripting tras varios intentos.")
            return False
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            logger.error("SAPGUI no es un objeto COM válido.")
            return False
        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            logger.error("No se pudo obtener ScriptingEngine de SAPGUI.")
            return False
        if amb == 'QAS':
            connection = application.OpenConnection("S/4 - QAS", True)
        elif amb == 'PRD':
            connection = application.OpenConnection("S/4 - PRD", True)
        else:
            logger.error(f"Ambiente desconocido: {amb}")
            return False
        if not type(connection) == win32com.client.CDispatch:
            logger.error("No se pudo abrir la conexión SAP.")
            return False
        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            logger.error("No se pudo obtener la sesión SAP.")
            return False
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = u
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = c
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        logger.info("Login SAP realizado correctamente.")
        return True
    except Exception as e:
        logger.error(f"Error en ingresarsap: {e}")
        return False
    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


def cerrar_sap(session=None):
    """
    Cierra SAP. Si se pasa una sesión, intenta cerrar elegantemente.
    Si no funciona o no hay sesión, mata todos los procesos SAP.
    """
    logger = logging.getLogger(__name__)
    
    # Intento 1: Cerrar elegantemente si tenemos sesión
    if session:
        try:
            logger.info("Cerrando SAP elegantemente...")
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
            logger.info("SAP cerrado elegantemente.")
            return True
        except Exception as e:
            logger.warning(f"Error cerrando elegantemente: {e}")
 