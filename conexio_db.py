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
        data = [['Cliente retenido', 1], ['Llamado reprogramado', 1], ['Cliente no retenido', 1], ['No quiere escuchar', 1], ['Contacto con el asesor', 1], ['Apoyo del asesor al ejecutivo', 1], ['Pendiente respuesta cliente', 1], ['Carta de revocación pendiente', 1], ['Contacto por correo', 1], ['Campaña exitosa', 1], ['Solicita renuncia', 1], ['Cliente desconoce venta', 1], ['No se pudo instalar mandato', 1], ['Anulado por cambio de producto Metlife', 1], ['Cliente vive en el extranjero', 1], ['Cliente activa mandato', 1], ['Queda vigente sin pagar', 1], ['Lo está viendo con Asesor', 1], ['Numero invalido', 0], ['Sin respuesta', 0], ['Buzón de voz', 0], ['Pagos al día', 0], ['Teléfono ocupado', 0], ['Teléfono apagado', 0], ['Campaña completada con 5 intentos', 0], ['Número equivocado', 0], ['Sin gestión de cierre', 0], ['Sin teléfono registrado', 0], ['Cliente no actualizado', 0], ['Temporalmente fuera de servicio', 0], ['Plazo previsto del producto', 0]]
        db = conectorDB()
        cursor = db.cursor()
        sql = """INSERT INTO estadout_pro_reac (decripcion, estado_contacto) VALUES (?, ?);"""
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