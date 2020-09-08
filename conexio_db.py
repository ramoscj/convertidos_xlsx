import mysql.connector
import pyodbc
import datetime

def conectorDB():
    try:
        entorno = False
        if entorno:
            cnx = mysql.connector.connect(
                host="localhost",
                user="root",
                password="",
                database="icom"
                )
        else:
            servidor = 'SOMARCJ\SOMAR01'
            db = 'icom'
            usuario = 'sa'
            clave = 'testdb'
            cnx = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+servidor+';DATABASE='+db+';UID='+usuario+';PWD='+ clave)
        return cnx
    except Exception as e:
        raise Exception('Error al conectar DB - %s' % e)
