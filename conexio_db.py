import mysql.connector
import pyodbc 

def conectorDB():
    try:
        entorno = True
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

# db = conectorDB()
# cursor = db.cursor()
sql = """MERGE ejecutivos WITH (SERIALIZABLE) AS Target USING (VALUES(?,?)) AS Source (rut, nombre) ON
        Target.rut = Source.rut
        WHEN MATCHED THEN 
             UPDATE SET Target.nombre = Source.nombre
        WHEN NOT MATCHED BY TARGET THEN
             INSERT (rut, nombre, plataforma, fecha_ingreso) VALUES (?, ?, ?, ?);"""
# cursor.execute(sql, ('22983208', 'Carlos Javier', '22983208', 'Carlos ramos', 'CRO 1', '2020-01-01'))