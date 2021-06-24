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
        data = [['Terminado con Exito', 'Terminado con Exito'], ['Pendiente', 'Pendiente'], ['Terminado sin Exito', 'Terminado sin Exito'], ['Sin Gestion', 'Sin Gestion'], ['Cliente no vigente','No gestionable'], ['Cliente al día','No gestionable'], ['Lo va a pensar','Pendiente'], ['Solicita contacto por correo','Pendiente'], ['Pendiente de endoso','Pendiente'], ['Mantiene su producto','Terminado con exito'], ['Desiste el producto','Terminado sin exito'], ['Anulado por cambio de producto','Terminado sin exito'], ['Sin gestión por cierre','Terminado sin exito'], ['Término programado de producto','No gestionable'], ['Queda vigente sin pagar','Terminado sin exito'], ['Espera de carta de anulación','Terminado sin exito']]
        db = conectorDB()
        cursor = db.cursor()
        sql = """INSERT INTO reactiva_estados_retencion (estado_retencion, estado_retencion_valido) VALUES (?, ?);"""
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