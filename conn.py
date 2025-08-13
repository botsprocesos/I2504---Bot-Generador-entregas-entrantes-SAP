import os
import time
from dotenv import load_dotenv
import pandas as pd
import win32com.client
from win32com.client import GetObject
from hdbcli import dbapi
import pythoncom


def connection(ambiente):
    load_dotenv()
    if ambiente == 'QAS':
        host =  os.getenv("HOST_LAB")      
        password = os.getenv("PASS_LAB")
        port =  os.getenv("PORT_LAB")
        user = os.getenv("USER_LAB") 

    elif ambiente == 'PRD':
        host =  os.getenv("HOST_RISE")
        password = os.getenv("PASS_RISE")
        port =  os.getenv("PORT_RISE")
        user = os.getenv("USER_RISE")

    conn = dbapi.connect(address=host, port=port, user=user, password=password, sslValidateCertificate=False )
    cursor = conn.cursor()
    cursor.execute("SET SCHEMA SAPABAP1")
    return conn