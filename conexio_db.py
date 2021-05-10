import datetime
import os

import mysql.connector
import pyodbc

from config_xlsx import ACCESO_DB


def conectorDB():
    try:
        cnx = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + ACCESO_DB['SERVIDOR'] + ';DATABASE=' + ACCESO_DB['NOMBRE_DB'] + ';UID=' + ACCESO_DB['USUARIO'] + ';PWD=' + ACCESO_DB['CLAVE'])
        return cnx
    except Exception as e:
        raise Exception('Error al conectar DB - %s' % e)

# print(conectorDB())

def insertar():
    try:
        data = [['Mantiene su producto', 'Mantiene su producto'], ['Cliente al día', 'No gestionable'], ['Cliente no vigente', 'No gestionable'], ['Sin gestión por cierre', 'No gestionable'], ['Término programado de producto', 'No gestionable'], ['Anulado por cambio de producto', 'No gestionable'], ['Lo va a pensar', 'Pendiente'], ['Pendiente de endoso', 'Pendiente'], ['Solicita contacto por correo', 'Pendiente'], ['Realiza Activación PAC/PAT', 'Realiza Activación PAC/PAT'], ['Realiza pago en línea', 'Realiza pago en línea'], ['Desiste el producto', 'Terminado sin Exito'], ['Espera de carta de anulación', 'Terminado sin Exito'], ['Queda vigente sin pagar', 'Terminado sin Exito'], ['Pendiente', 'Pendiente'], ['Terminado sin Exito', 'Terminado sin Exito']]
        db = conectorDB()
        cursor = db.cursor()
        sql = """INSERT INTO proactiva_estados_retencion (estado_retencion, estado_retencion_valido) VALUES (?, ?);"""
        cursor.executemany(sql, data)
        db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar: %s' % (e))
    finally:
        cursor.close()
        db.close()

# print(insertar())