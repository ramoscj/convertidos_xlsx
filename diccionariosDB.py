from conexio_db import conectorDB

def buscarEjecutivosDb():
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT rut, nombre FROM ejecutivos"""
        cursor.execute(sql)
        for (rut, nombre) in cursor:
            ejecutivos[nombre] = {'RUT': rut, 'NOMBRE': nombre}
        return ejecutivos
    except Exception as e:
        raise Exception('Error buscarEjecutivosDb: %s' % e)
    finally:
        cursor.close()
        db.close()

def buscarCamphnasDb():
    try:
        db = conectorDB()
        cursor = db.cursor()
        campanhas = dict()
        sql = """SELECT codigo, nombre FROM codigos_cro"""
        cursor.execute(sql)
        for (codigo, nombre) in cursor:
            campanhas[nombre] = {'CODIGO': codigo, 'NOMBRE': nombre}
        return campanhas
    except Exception as e:
        raise Exception('Error buscarCamphnasDb: %s' % e)
    finally:
        cursor.close()
        db.close()

def buscarRutEjecutivosDb():
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT rut, nombre FROM ejecutivos"""
        cursor.execute(sql)
        for (rut, nombre) in cursor:
            ejecutivos[rut] = {'RUT': rut, 'NOMBRE': nombre}
        return ejecutivos
    except Exception as e:
        raise Exception('Error buscarRutEjecutivosDb: %s' % e)
    finally:
        cursor.close()
        db.close()