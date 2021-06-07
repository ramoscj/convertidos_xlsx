from conexio_db import conectorDB
from validaciones_texto import separarNombreApellido


def buscarEjecutivosDb():
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT id, id_empleado, plataforma FROM ejecutivos"""
        cursor.execute(sql)
        for (id, id_empleado, plataforma) in cursor:
            ejecutivos[id_empleado] = {'ID': id, 'ID_EMEPLEADO': id_empleado, 'PLATAFORMA': plataforma}
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

def buscarRutEjecutivosDb(ultimoDiaMes, primerDiaMes):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivos = dict()
        sql = """SELECT id_empleado, plataforma, fecha_ingreso, fecha_desvinculacion FROM ejecutivos WHERE isnull(fecha_desvinculacion, ?) >= ?"""
        cursor.execute(sql, (ultimoDiaMes, primerDiaMes))
        for (id_empleado, plataforma, fecha_ingreso, fecha_desvinculacion) in cursor:
            if fecha_desvinculacion is not None:
                fecha_desvinculacion = fecha_desvinculacion.strftime("%d-%m-%Y")
            ejecutivos[id_empleado] = {'ID_EMPLEADO': id_empleado, 'PLATAFORMA': plataforma, 'FECHA_INGRESO': fecha_ingreso.strftime("%d-%m-%Y"), 'FECHA_DESVINCULACION': fecha_desvinculacion}
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
        sql = """SELECT id_ejecutivo, numero_poliza, campana_id, nombre_campana, cobranza_pro, pacpat_pro, cobranza_rel_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, periodo, fecha_reliquidacion, fecha_cierre, numero_poliza_certificado FROM proactiva_campanas_ejecutivos LEFT JOIN proactiva_campanas_periodo_ejecutivos pcpe ON proactiva_campanas_ejecutivos.id_periodo_ejecutivo = pcpe.id WHERE pcpe.periodo = ? AND reliquidacion = 1 AND fecha_reliquidacion is NULL"""
        cursor.execute(sql, (mesAnterior))
        for (id_ejecutivo, numero_poliza, campana_id, nombre_campana, cobranza_pro, pacpat_pro, cobranza_rel_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, periodo, fecha_reliquidacion, fecha_cierre, numero_poliza_certificado) in cursor:
            polizasParaRequilidar[numero_poliza] = {'COBRANZA_RL_PRO': cobranza_rel_pro, 'PACPAT_RL_PRO': pacpat_rel_pro, 'ESTADO_PRO': estado_pro, 'ESTADO_UT_PRO': estado_ut_pro, 'CODIGO_EMPLEADO': id_ejecutivo, 'NOMBRE_CAMPANA': nombre_campana, 'CAMPAÑA_ID': campana_id, 'POLIZA': numero_poliza, 'FECHA_CIERRE': fecha_cierre, 'NUMERO_POLIZA_CERTIFICADO': numero_poliza_certificado}
        return polizasParaRequilidar
    except Exception as e:
        raise Exception('Error buscarPolizasReliquidar: %s' % e)
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
        sql = """SELECT count(*) FROM proactiva_campanas_ejecutivos LEFT JOIN proactiva_campanas_periodo_ejecutivos pcpe ON proactiva_campanas_ejecutivos.id_periodo_ejecutivo = pcpe.id
        WHERE pcpe.periodo = ?"""
        cantidadCampanas = cursor.execute(sql, (fechaPeriodo)).fetchone()
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
        raise Exception('Error buscarEjecutivosDb2(): %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoUtContacto():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaEstadoUt = dict()
        sql = """SELECT descripcion FROM estadout_pro_reac WHERE estado_contacto = 1"""
        cursor.execute(sql)
        for (descripcion,) in cursor:
            listaEstadoUt.setdefault(descripcion, 1)
        return listaEstadoUt
    except Exception as e:
        raise Exception('Error listaEstadoUtContacto(): %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoUtNoContacto():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaEstadoUt = dict()
        sql = """SELECT descripcion FROM estadout_pro_reac WHERE estado_contacto = 0"""
        cursor.execute(sql)
        for (descripcion,) in cursor:
            listaEstadoUt.setdefault(descripcion, 1)
        return listaEstadoUt
    except Exception as e:
        raise Exception('Error listaEstadoUtNoContacto: %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoUtAll():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaEstadoUt = dict()
        sql = """SELECT descripcion, id FROM estadout_pro_reac order by id"""
        cursor.execute(sql)
        for (descripcion, id) in cursor:
            listaEstadoUt.setdefault(descripcion, id)
        return listaEstadoUt
    except Exception as e:
        raise Exception('Error def listaEstadoUtAll(): %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoUtDesc():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaEstadoUt = dict()
        sql = """SELECT id, descripcion FROM estadout_pro_reac order by id"""
        cursor.execute(sql)
        for (id, descripcion) in cursor:
            listaEstadoUt.setdefault(id, descripcion)
        return listaEstadoUt
    except Exception as e:
        raise Exception('Error def listaEstadoUtDesc(): %s' % e)
    finally:
        cursor.close()
        db.close()

def ReliquidacionesPorPeriodo(fechaPeriodo):
    try:
        db = conectorDB()
        cursor = db.cursor()

        sql = """SELECT count(*) FROM retenciones_por_reliquidar WHERE fecha_proceso = ?"""
        cantidadCampanas = cursor.execute(sql, (fechaPeriodo)).fetchone()
        return cantidadCampanas[0]
    except Exception as e:
        raise Exception('Error buscar Polizas ReliquidacionesPorPeriodo(): %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoRetencionProactiva():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaSalida = dict()
        sql = """SELECT estado_retencion, id FROM proactiva_estados_retencion order by id"""
        cursor.execute(sql)
        for (estado_retencion, id) in cursor:
            listaSalida.setdefault(estado_retencion, id)
        return listaSalida
    except Exception as e:
        raise Exception('Error def listaEstadoRetencionProactiva(): %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoRetencionDesc():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaSalida = dict()
        sql = """SELECT id, estado_retencion FROM proactiva_estados_retencion order by id"""
        cursor.execute(sql)
        for (id, estado_retencion) in cursor:
            listaSalida.setdefault(id, estado_retencion)
        return listaSalida
    except Exception as e:
        raise Exception('Error def listaEstadoRetencionDesc(): %s' % e)
    finally:
        cursor.close()
        db.close()

def listaEstadoUtCro():
    try:
        db = conectorDB()
        cursor = db.cursor()
        listaEstadoUt = dict()
        sql = """SELECT descripcion, id FROM estadout_cro order by id"""
        cursor.execute(sql)
        for (descripcion, id) in cursor:
            listaEstadoUt.setdefault(str(descripcion).upper(), id)
        return listaEstadoUt
    except Exception as e:
        raise Exception('Error def listaEstadoUtCro(): %s' % e)
    finally:
        cursor.close()
        db.close()

# x = ReliquidacionesPorPeriodo('01/12/2020')
# x = listaEstadoRetencionDesc()
# print(listaEstadoUtCro())