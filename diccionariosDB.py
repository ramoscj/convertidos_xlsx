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

def buscarPolizasReliquidar(mesAnterior):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaRequilidar = dict()
        sql = """SELECT rut_ejecutivo, id_cliente, numero_poliza, campana_id, cobranza_pro, cobranza_rel_pro, pacpat_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre FROM retenciones_por_reliquidar WHERE fecha_proceso = ?"""
        cursor.execute(sql, (mesAnterior))
        for (rut_ejecutivo, id_cliente, numero_poliza, campana_id, cobranza_pro, cobranza_rel_pro, pacpat_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre) in cursor:
            polizasParaRequilidar[numero_poliza] = {'COBRANZA_PRO': cobranza_pro, 'COBRANZA_REL_PRO': cobranza_rel_pro, 'PACPAT_PRO': pacpat_pro, 'PACPAT_REL_PRO': pacpat_rel_pro, 'ESTADO_PRO': estado_pro, 'ESTADO_UT_PRO': estado_ut_pro, 'RUT': rut_ejecutivo, 'CAMPAÑA_ID': campana_id, 'POLIZA': numero_poliza, 'ID_CLIENTE': id_cliente, 'FECHA_CIERRE': fecha_cierre}
        return polizasParaRequilidar
    except Exception as e:
        raise Exception('Error buscar Polizas para reliquidar: %s' % e)
    finally:
        cursor.close()
        db.close()

def buscarPolizasReliquidarAll():
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaRequilidar = dict()
        sql = """SELECT rut_ejecutivo, id_cliente, numero_poliza FROM retenciones_por_reliquidar"""
        cursor.execute(sql)
        for (rut_ejecutivo, id_cliente, numero_poliza) in cursor:
            pk = '%s_%s_%s' % (id_cliente, rut_ejecutivo, numero_poliza)
            polizasParaRequilidar[pk] = {'RUT': rut_ejecutivo, 'POLIZA': numero_poliza, 'ID_CLIENTE': id_cliente}
        return polizasParaRequilidar
    except Exception as e:
        raise Exception('Error buscar Polizas para reliquidarAll: %s' % e)
    finally:
        cursor.close()
        db.close()

# print(buscarPolizasReliquidarAll())
