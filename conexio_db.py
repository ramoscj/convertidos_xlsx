import mysql.connector

def conectorDB():
    try:
        db = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="ICOM"
            )
        return db
    except Exception as e:
        raise Exception('Error al conectar DB - %s' % e)