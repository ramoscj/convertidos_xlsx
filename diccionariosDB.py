from conexio_db import conectorDB
from validaciones_texto import separarNombreApellido

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
        sql = """SELECT rut, nombre, nombre_rrh, plataforma FROM ejecutivos"""
        cursor.execute(sql)
        for (rut, nombre, nombre_rrh, plataforma) in cursor:
            ejecutivos[rut] = {'RUT': rut, 'NOMBRE': nombre, 'PLATAFORMA': ''.join((plataforma).split()), 'NOMBRE_RRH': nombre_rrh}
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
        sql = """SELECT rut, nombre_rrh, plataforma, fecha_ingreso, fecha_desvinculacion FROM ejecutivos WHERE isnull(fecha_desvinculacion, ?) >= ? AND plataforma <> 'CORET REACTIVA' AND plataforma <> 'CORET PROACTIVA'"""
        cursor.execute(sql, (ultimoDiaMes, primerDiaMes))
        for (rut, nombre_rrh, plataforma, fecha_ingreso, fecha_desvinculacion) in cursor:
            if fecha_desvinculacion is not None:
                fecha_desvinculacion = fecha_desvinculacion.strftime("%d-%m-%Y")
            apellidoPaterno, apellidoMaterno, nombres = separarNombreApellido(nombre_rrh)
            ejecutivos[rut] = {'RUT': rut, 'APELLIDO_PATERNO': apellidoPaterno, 'APELLIDO_MATERNO': apellidoMaterno, 'NOMBRES': nombres, 'PLATAFORMA': plataforma, 'FECHA_INGRESO': fecha_ingreso.strftime("%d-%m-%Y"), 'FECHA_DESVINCULACION': fecha_desvinculacion}
        return ejecutivos
    except Exception as e:
        raise Exception('Error buscarEjecutivosAllDb: %s' % e)
    finally:
        cursor.close()
        db.close()

# print(buscarRutEjecutivosDb())
