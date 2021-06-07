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
        data = [['Campaña exitosa',1], ['Teléfono ocupado',0], ['Sin respuesta',0], ['Campaña completada con 5 intentos',0], ['Buzón voz',0], ['Llamado reprogramado',1], ['Contacto por correo',1], ['Teléfono apagado',0], ['Número equivocado',0], ['Numero invalido',0], ['Solicita renuncia',1], ['No quiere escuchar',1], ['Cliente desconoce venta',1], ['Temporalmente fuera de servicio',0], ['Cliente vive en el extranjero',0], ['Sin teléfono registrado',0], ['Cliente no retenido',1], ['No contesta',0], ['Pendiente de envío de póliza',1], ['Desconoce venta',1], ['Pagos al día',0], ['No se pudo instalar mandato',1], ['Lo está viendo con asesor',1], ['Queda vigente sin pagar',1], ['Cliente activa mandato',1], ['Sin gestión por cierre',0], ['Anulado por cambio de producto Metlife',0], ['Pendiente de respuesta del cliente',1], ['Pendiente de carta de revocación',1], ['Cliente retenido',1], ['Apoyo de asesor para retención',1], ['Buzón voz',0], ['No iniciado',0], ['No iniciado',0], ['Completado',1], ['Sin iniciar',0], ['Buzón de voz',0], ['Sin Gestion',0], ['Contacto con el asesor',1], ['Pendiente respuesta cliente',1], ['Cliente no actualizado',0], ['Sin gestión de cierre',0], ['Plazo previsto del producto',1]]
        db = conectorDB()
        cursor = db.cursor()
        sql = """INSERT INTO estadout_cro (descripcion, contactado) VALUES (?, ?);"""
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