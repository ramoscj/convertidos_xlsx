import mysql.connector
import pyodbc
import datetime
import os

from config_xlsx import ACCESO_DB

def conectorDB():
    try:
        cnx = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + ACCESO_DB['SERVIDOR'] + ';DATABASE=' + ACCESO_DB['NOMBRE_DB'] + ';UID=' + ACCESO_DB['USUARIO'] + ';PWD=' + ACCESO_DB['CLAVE'])
        return cnx
    except Exception as e:
        raise Exception('Error al conectar DB - %s' % e)

# print(conectorDB())