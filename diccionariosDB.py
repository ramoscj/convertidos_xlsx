from conexio_db import conectorDB
from validaciones_texto import separarNombreApellido


def buscarEjecutivosDb():
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT id, rut, nombre FROM ejecutivos"""
        cursor.execute(sql)
        for (id, rut, nombre) in cursor:
            ejecutivos[nombre] = {'ID': id, 'RUT': rut, 'NOMBRE': nombre}
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

def buscarEjecutivosVinculados(ultimoDiaMes, primerDiaMes):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT rut, nombre_rrh, plataforma, fecha_ingreso, fecha_desvinculacion FROM ejecutivos WHERE isnull(fecha_desvinculacion, ?) >= ?"""
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

def buscarPolizasReliquidar(mesAnterior):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaRequilidar = dict()
        sql = """SELECT codigo_empleado, id_cliente, numero_poliza, campana_id, nombre_campana, cobranza_pro, pacpat_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre FROM retenciones_por_reliquidar WHERE fecha_proceso = ?"""
        cursor.execute(sql, (mesAnterior))
        for (codigo_empleado, id_cliente, numero_poliza, campana_id, nombre_campana, cobranza_pro, pacpat_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre) in cursor:
            polizasParaRequilidar[numero_poliza] = {'COBRANZA_PRO': cobranza_pro, 'PACPAT_PRO': pacpat_pro, 'ESTADO_PRO': estado_pro, 'ESTADO_UT_PRO': estado_ut_pro, 'CODIGO_EMPLEADO': codigo_empleado, 'NOMBRE_CAMPANA': nombre_campana, 'CAMPAÑA_ID': campana_id, 'POLIZA': numero_poliza, 'ID_CLIENTE': id_cliente, 'FECHA_CIERRE': fecha_cierre}
        return polizasParaRequilidar
    except Exception as e:
        raise Exception('Error buscar Polizas para buscarPolizasReliquidar: %s' % e)
    finally:
        cursor.close()
        db.close()

def buscarPolizasReliquidarAll():
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaRequilidar = dict()
        sql = """SELECT codigo_empleado, id_cliente, numero_poliza FROM retenciones_por_reliquidar"""
        cursor.execute(sql)
        for (codigo_empleado, id_cliente, numero_poliza) in cursor:
            pk = '%s_%s_%s' % (id_cliente, codigo_empleado, numero_poliza)
            polizasParaRequilidar[pk] = {'RUT': codigo_empleado, 'POLIZA': numero_poliza, 'ID_CLIENTE': id_cliente}
        return polizasParaRequilidar
    except Exception as e:
        raise Exception('Error buscar Polizas para buscarPolizasReliquidarAll: %s' % e)
    finally:
        cursor.close()
        db.close()

def periodoCampanasEjecutivos(fechaPeriodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        campanasEjecutivos = dict()
        sql = """SELECT id, id_ejecutivo, periodo FROM proactiva_campanas_periodo_ejecutivos WHERE periodo = ?"""
        cursor.execute(sql, (fechaPeriodo))
        for (id, id_ejecutivo, periodo) in cursor:
            campanasEjecutivos[id_ejecutivo] = {'ID': id, 'ID_EMPLEADO': id_ejecutivo, 'PERIODO': periodo}
        return campanasEjecutivos
    except Exception as e:
        raise Exception('Error buscar Polizas para periodoCampanasEjecutivos: %s' % e)
    finally:
        cursor.close()
        db.close()

def CamapanasPorPeriodo(fechaPeriodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        idEjecutivos = []
        ejecutivosExistentes = periodoCampanasEjecutivos(fechaPeriodo)
        for valores in ejecutivosExistentes.values():
            idEjecutivos.append(valores['ID'])

        cantidadRegistros = ', '.join('?' * len(idEjecutivos))
        sql = """SELECT count(*) FROM proactiva_campanas_ejecutivos WHERE id_periodo_ejecutivo IN ({valores})""".format(valores=cantidadRegistros)
        cantidadCampanas = cursor.execute(sql, (idEjecutivos)).fetchone()
        return cantidadCampanas[0]
    except Exception as e:
        raise Exception('Error buscar Camapañas CamapanasPorPeriodo: %s' % e)
    finally:
        cursor.close()
        db.close()

def buscarEjecutivosDb2():
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT id, id_empleado, plataforma FROM ejecutivos_2"""
        cursor.execute(sql)
        for (id, idEmpleado, plataforma) in cursor:
            ejecutivos[idEmpleado] = {'ID': id, 'CODIGO_EMPLEADO': idEmpleado, 'PLATAFORMA': plataforma}
        return ejecutivos
    except Exception as e:
        raise Exception('Error buscarEjecutivosDb2: %s' % e)
    finally:
        cursor.close()
        db.close()

# print(CamapanasPorPeriodo('01/12/2020'))
