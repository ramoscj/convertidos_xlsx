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
        sql = """SELECT nombre FROM codigos_cro"""
        cursor.execute(sql)
        for (nombre,) in cursor:
            campanhas[nombre] = {'NOMBRE': nombre}
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
        sql = """SELECT rut, nombre, plataforma FROM ejecutivos"""
        cursor.execute(sql)
        for (rut, nombre, plataforma) in cursor:
            ejecutivos[rut] = {'RUT': rut, 'NOMBRE': nombre, 'PLATAFORMA': ''.join((plataforma).split())}
        return ejecutivos
    except Exception as e:
        raise Exception('Error buscarRutEjecutivosDb: %s' % e)
    finally:
        cursor.close()
        db.close()

def buscarEjecutivosAllDb(ultimoDiaMes, primerDiaMes):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT rut, nombre, plataforma, fecha_ingreso, fecha_desvinculacion FROM ejecutivos WHERE isnull(fecha_desvinculacion, ?) >= ?"""
        cursor.execute(sql, (ultimoDiaMes, primerDiaMes))
        for (rut, nombre, plataforma, fecha_ingreso, fecha_desvinculacion) in cursor:
            if fecha_desvinculacion is not None:
                fecha_desvinculacion = fecha_desvinculacion.strftime("%d-%m-%Y")
            ejecutivos[rut] = {'RUT': rut, 'NOMBRE': nombre, 'PLATAFORMA': plataforma, 'FECHA_INGRESO': fecha_ingreso.strftime("%d-%m-%Y"), 'FECHA_DESVINCULACION': fecha_desvinculacion}
        return ejecutivos
    except Exception as e:
        raise Exception('Error buscarEjecutivosAllDb: %s' % e)
    finally:
        cursor.close()
        db.close()

# print(buscarRutEjecutivosDb())