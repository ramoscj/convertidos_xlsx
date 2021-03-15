import datetime
import contextlib
import os

from openpyxl import load_workbook
from tqdm import tqdm

from complementoCliente import (LOG_COMPLEMENTO_CLIENTE,
                                extraerComplementoCliente)
from conexio_db import conectorDB
from config_xlsx import PATH_XLSX, PROACTIVA_CONFIG_XLSX, listaEstadoContactado
from diccionariosDB import (buscarEjecutivosDb2, buscarPolizasReliquidar,
                            buscarPolizasReliquidarAll, periodoCampanasEjecutivos, CamapanasPorPeriodo)
from escribir_txt import (salidaArchivoTxt, salidaArchivoTxtProactiva,
                          salidaLogTxt, salidaInsertBulkCampanas)
from validaciones_texto import (convertirALista,
                                formatearFechaMesAnterior, formatearIdCliente,
                                formatearNumeroPoliza, mesSiguienteUltimoDia,
                                primerDiaMes, setearCelda, setearCelda2,
                                setearFechaCelda, ultimoDiaMes,
                                validarEncabezadoXlsx,
                                convertirListaCampana,
                                setearCampanasPorEjecutivo)

LOG_PROCESO_PROACTIVA = dict()
polizasNoAprobadas = dict()
campanasPorEjecutivos = dict()

def getEstado(celdaEstado):
    listaContactado = {'Terminado con Exito': 1 , 'Pendiente': 2 , 'Terminado sin Exito': 3}
    if listaContactado.get(str(celdaEstado.value)):
        return listaContactado[celdaEstado.value]
    elif str(celdaEstado.value) == 'Sin Gestion':
        return 0
    else:
        return False

def getEstadoUt(celdaEstadoUt):
    estadoUt = celdaEstadoUt
    listaEstadoUt = PROACTIVA_CONFIG_XLSX['LISTA_ULTIMA_TAREA']
    if listaEstadoUt.get(estadoUt.value):
        return listaEstadoUt[estadoUt.value]
    elif estadoUt.value is None:
        return 0
    else:
        return False

def estadoCertificadoPoliza(numeroPoliza):
    resto, separador, nroCertificado = str(numeroPoliza).partition("_")
    if str(nroCertificado):
        return nroCertificado
    return str(0)

def aprobarCobranza(nroPolizaCertificado, fechaCierre, nroPolizaCliente, fecUltimoPago):
    if fecUltimoPago is not None and fechaCierre is not None:
        if nroPolizaCliente == nroPolizaCertificado and  fecUltimoPago >= fechaCierre:
            return 1
    elif nroPolizaCliente == nroPolizaCertificado and fecUltimoPago is None:
        return 1
    return 0

def aprobarActivacion(estadoMandato, fechaMandato, fechaCierre):
    if estadoMandato == 'APROBADO ENTIDAD RECAUDADORA':
        if fechaMandato is None:
            return 1
        elif fechaMandato is not None and fechaCierre is not None:
            if fechaMandato >= fechaCierre:
                return 1
    return 0

def insertarPolizaNoAprobada(dataPolizas:list):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasInsertar = convertirALista(dataPolizas)
        sql = """INSERT INTO retenciones_por_reliquidar (codigo_empleado, id_cliente, numero_poliza, campana_id, nombre_campana, cobranza_pro, cobranza_rel_pro, pacpat_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?);"""
        cursor.executemany(sql, polizasInsertar)
        db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar polizas para reliquidar: %s' % (e))
    finally:
        cursor.close()
        db.close()

def insertarPeriodoCampanaEjecutivos(camapanasEjecutivos: dict, fechaProceso):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivosExistentes = periodoCampanasEjecutivos(fechaProceso)
        periodoEjecutivos = convertirListaCampana(camapanasEjecutivos, ejecutivosExistentes, fechaProceso)
        if len(periodoEjecutivos) > 0:
            sql = """INSERT INTO proactiva_campanas_periodo_ejecutivos (id_ejecutivo, periodo) VALUES (?, ?);"""
            cursor.executemany(sql, periodoEjecutivos)
            db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar Periodo de Ejecutivos: %s' % (e))
    finally:
        cursor.close()
        db.close()

def insertarCampanaEjecutivos(camapanasEjecutivos: dict, fechaProceso):
    try:
        db = conectorDB()
        camapanasPeriodoEjecutivos = periodoCampanasEjecutivos(fechaProceso)
        campanasPorPeriodo = []
        cursor = db.cursor()

        for valores in camapanasEjecutivos.values():
            idEjecutivo = valores['ID_EJECUTIVO']
            if camapanasPeriodoEjecutivos.get(idEjecutivo):
                campanasPorPeriodo += setearCampanasPorEjecutivo(valores['CAMPANAS'], camapanasPeriodoEjecutivos[idEjecutivo]['ID'])

        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'INSERTAR_CAMPANAS_EJECUTIVOS': '-----------------------------------------------------' })
        campanasExistentes = CamapanasPorPeriodo(fechaProceso)
        if limpiarTablaCamapanasEjecutivos(fechaProceso):
            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'LIMPIAR_CAMAPAÑAS_EJECUTIVOS': 'InsertarCampanaEjecutivos;Se eliminaron {campanas} camapañas existentes'.format(campanas= campanasExistentes)})

        sql = """INSERT INTO proactiva_campanas_ejecutivos (id_periodo_ejecutivo, nombre_cliente, fecha_creacion, nombre_campana, numero_poliza, fecha_cierre, estado_retencion, estado_ut) VALUES (?, ?, ?, ?, ?, ?, ?, ?);"""
        cursor.executemany(sql, campanasPorPeriodo)
        db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar Campañas de Ejecutivos del Periodo: %s' % (e))
    finally:
        cursor.close()
        db.close()

def limpiarTablaCamapanasEjecutivos(fechaProceso):
    try:
        db = conectorDB()
        cursor = db.cursor()
        idEjecutivos = []


        ejecutivosExistentes = periodoCampanasEjecutivos(fechaProceso)
        for valores in ejecutivosExistentes.values():
            idEjecutivos.append([valores['ID']])

        sql = """DELETE FROM proactiva_campanas_ejecutivos WHERE id_periodo_ejecutivo = ?;"""
        cursor.executemany(sql, idEjecutivos)
        db.commit()
        return cursor
    except Exception as e:
        db.rollback()
        raise Exception('Error al eliminar Campañas de Ejecutivos existentes: %s' % (e))
    finally:
        cursor.close()
        db.close()

def agregarCampanasPorEjecutivo(idEmpleado: int, valoresCampanas: []):

    if campanasPorEjecutivos.get(idEmpleado):
        campanasPorEjecutivos[idEmpleado]['CAMPANAS'].append(valoresCampanas)
    else:
        campanasPorEjecutivos[idEmpleado] = {'ID_EJECUTIVO': idEmpleado, 'CAMPANAS': [valoresCampanas]}
    return 1

def polizasReliquidadas(periodo, complementoCliente):
    mesAnterior = formatearFechaMesAnterior(periodo)
    fechaIncioMes = primerDiaMes(periodo)
    polizasParaReliquidar = buscarPolizasReliquidar(mesAnterior)
    polizasAprobadaReliquidar = dict()

    for poliza in polizasParaReliquidar.values():
        numeroPolizaCertificado = estadoCertificadoPoliza(poliza['POLIZA'])
        numeroPoliza = formatearNumeroPoliza(poliza['POLIZA'])
        nombreCampana = poliza['NOMBRE_CAMPANA']
        fechaCierre = poliza['FECHA_CIERRE']
        numeroPolizaCliente = complementoCliente[numeroPoliza]['NRO_CERT']
        fecUltimoPago = complementoCliente[numeroPoliza]['FEC_ULT_PAG']
        cobranzaRelPro = 0
        pacpatRelPro = 0
        if poliza['COBRANZA_PRO'] > 0:
            if complementoCliente.get(numeroPoliza):
                cobranzaRelPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, complementoCliente[numeroPoliza]['NRO_CERT'] , complementoCliente[numeroPoliza]['FEC_ULT_PAG'])
                if cobranzaRelPro == 0:
                    mensaje = 'PolizaReliquidacion;No cumple condicion de retencion COBRO para Reliquidacion;%s' % (numeroPoliza)
                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA_RELIQUIDACION': mensaje})

        if poliza['PACPAT_PRO'] > 0:
            estadoMandato = complementoCliente[numeroPoliza]['ESTADO_MANDATO']
            fecMandato = complementoCliente[numeroPoliza]['FECHA_MANDATO']
            if estadoMandato is not None:
                pacpatRelPro = aprobarActivacion(estadoMandato, fecMandato, fechaCierre)
            else:
                pacpatRelPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
            if pacpatRelPro == 0:
                mensaje = 'PolizaReliquidacion;No cumple condicion de retencion ACTIVACION para Reliquidacion;%s' % (numeroPoliza)
                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_ACTIVACION_RELIQUIDACION': mensaje})

        if cobranzaRelPro > 0 or pacpatRelPro > 0:
            polizasAprobadaReliquidar[numeroPoliza] = {'COBRANZA_PRO': 0, 'COBRANZA_REL_PRO': cobranzaRelPro, 'PACPAT_PRO': 0, 'PACPAT_REL_PRO': pacpatRelPro, 'ESTADO_PRO': poliza['ESTADO_PRO'], 'ESTADO_UT_PRO': poliza['ESTADO_UT_PRO'], 'CODIGO_EMPLEADO': poliza['CODIGO_EMPLEADO'], 'CAMPAÑA_ID': poliza['CAMPAÑA_ID'], 'NOMBRE_CAMAPANA': nombreCampana, 'POLIZA': numeroPoliza, 'ID_CLIENTE': poliza['ID_CLIENTE']}
    
    if len(polizasAprobadaReliquidar) > 0:
        actualizarPolizasReliquidadas(polizasAprobadaReliquidar, fechaIncioMes)
        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZAS_RELIQUIDADS': 'Existen {polizas} polizas que se van reliquidar'.format(polizas=len(polizasAprobadaReliquidar))})
    else:
        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZAS_RELIQUIDADS': 'No Existen polizas para reliquidar del periodo: {mesProceso}'.format(mesProceso=mesAnterior)})

    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZAS_RELIQUIDADS': '-----------------------------------------------------' })
    return polizasAprobadaReliquidar

def actualizarPolizasReliquidadas(polizasReliquidadas, fechaProceso):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaActualizar = []

        for valores in polizasReliquidadas.values():
            polizasParaActualizar.append([fechaProceso, valores['POLIZA']])

        sql = """UPDATE retenciones_por_reliquidar SET fecha_reliquidacion = ? WHERE numero_poliza = ? AND fecha_reliquidacion IS NULL;"""
        cursor.executemany(sql,polizasParaActualizar)
        db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar polizas para reliquidar: %s' % (e))
    finally:
        cursor.close()
        db.close()

def validarRetencionesPolizas(valoresEntrada: dict, complementoCliente: dict):

    estadoRetencion = valoresEntrada['ESTADO_RETENCION']
    nombreCampana = valoresEntrada['NOMBRE_CAMPAÑA']
    numeroPoliza = valoresEntrada['NUMERO_POLIZA']
    idEmpleado = valoresEntrada['ID_EMPLEADO']
    numeroPolizaOriginal = valoresEntrada['NUMER_POLIZA_ORIGINAL']
    fechaCierre = valoresEntrada['FECHA_CIERRE']
    campanaId = valoresEntrada['CAMPAÑA_ID']
    estadoValido = valoresEntrada['ESTADO_VALIDO']
    estadoUtValido = valoresEntrada['ESTADO_VALIDOUT']
    fechaIncioMes = valoresEntrada['FECHA_INICIO_MES']
    celdaNroPoliza = valoresEntrada['CELDA_NROPOLIZA']
    idCiente = valoresEntrada['ID_CLIENTE']
    cobranzaPro = 0
    pacpatPro = 0
    listaConsiderarRetencion = {'Mantiene su producto': 1, 'Realiza pago en línea': 2, 'Realiza Activación PAC/PAT': 3}
    polizasReliquidarDb = buscarPolizasReliquidarAll()

    retencion = listaConsiderarRetencion.get(estadoRetencion)
    if nombreCampana == 'CO RET - Cobranza' and retencion == 3:
        cobranzaPro = 1
        pacpatPro = 1
    elif nombreCampana == 'CO RET - Fallo en Instalacion de Mandato' and retencion == 1:
        cobranzaPro = 1
        pacpatPro = 1
    elif nombreCampana == 'CO RET - Cobranza' and retencion == 1 or nombreCampana == 'CO RET - Cobranza' and retencion == 2:
        cobranzaPro = 1
    elif nombreCampana == 'CO RET - Fallo en Instalacion de Mandato' and retencion == 2:
        cobranzaPro = 1
    elif nombreCampana == 'CO RET - Fallo en Instalacion de Mandato' and retencion == 3:
        pacpatPro = 1


    numeroPolizaCertificado = estadoCertificadoPoliza(numeroPoliza)
    fecUltimoPago = complementoCliente[numeroPoliza]['FEC_ULT_PAG']
    numeroPolizaCliente = complementoCliente[numeroPoliza]['NRO_CERT']
    pk2 = '%s_%s_%s' % (idCiente, idEmpleado, numeroPolizaOriginal)
    if cobranzaPro > 0:
        cobranzaPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
        if cobranzaPro == 0:
            if not polizasReliquidarDb.get(pk2):
                if not polizasNoAprobadas.get(pk2):
                    polizasNoAprobadas[pk2] = {'ID_EMPLEADO': idEmpleado, 'ID_CLIENTE': idCiente, 'NRO_POLIZA': numeroPolizaOriginal, 'ID_CAMPAÑA': campanaId, 'NOMBRE_CAMPANA': nombreCampana, 'COBRANZA_PRO': 1, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': pacpatPro, 'PACPAT_REL_PRO': 0, 'ESTADO_VALIDO': estadoValido, 'ESTADO_VALIDOUT': estadoUtValido, 'FECHA_INICIO_MES': fechaIncioMes, 'FECHA_CIERRE': fechaCierre}
            else:
                celdaCoordenada = setearCelda2(celdaNroPoliza, 0)
                mensaje = '%s;Poliza para reliquidar esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA': mensaje})
        else:
            celdaCoordenada = setearCelda2(celdaNroPoliza, 0)
            mensaje = '%s;Poliza no cumple condicion de retencion COBRO;%s' % (celdaCoordenada, numeroPoliza)
            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA': mensaje})

    if pacpatPro > 0:
        estadoMandato = complementoCliente[numeroPoliza]['ESTADO_MANDATO']
        fecMandato = complementoCliente[numeroPoliza]['FECHA_MANDATO']
        if estadoMandato is not None:
            pacpatPro = aprobarActivacion(estadoMandato, fecMandato, fechaCierre)
        else:
            pacpatPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
        if pacpatPro == 0:
            if not polizasReliquidarDb.get(pk2):
                if not polizasNoAprobadas.get(pk2):
                    polizasNoAprobadas[pk2] = {'ID_EMPLEADO': idEmpleado, 'ID_CLIENTE': idCiente, 'NRO_POLIZA': numeroPolizaOriginal, 'ID_CAMPAÑA': campanaId, 'NOMBRE_CAMPANA': nombreCampana, 'COBRANZA_PRO': cobranzaPro, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': 1, 'PACPAT_REL_PRO': 0, 'ESTADO_VALIDO': estadoValido, 'ESTADO_VALIDOUT': estadoUtValido, 'FECHA_INICIO_MES': fechaIncioMes, 'FECHA_CIERRE': fechaCierre}
                else:
                    polizasNoAprobadas[pk2]['PACPAT_PRO'] = 1
            else:
                celdaCoordenada = setearCelda2(celdaNroPoliza, 0)
                mensaje = '%s;Poliza para reliquidar esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_ACTIVACION': mensaje})
        else:
            celdaCoordenada = setearCelda2(celdaNroPoliza, 0)
            mensaje = '%s;Poliza no cumple condicion de retencion ACTIVACION;%s' % (celdaCoordenada, numeroPoliza)
            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA': mensaje})

    return cobranzaPro, pacpatPro

def leerArchivoProactiva(archivoEntrada, periodo):
    try:
        encabezadoXls = PROACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = PROACTIVA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = PROACTIVA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        coordenadaEcabezado = PROACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        archivo_correcto = validarEncabezadoXlsx(hoja[coordenadaEcabezado], encabezadoXls, archivoEntrada)
        if type(archivo_correcto) is not dict:
            filaSalidaXls = dict()

            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            fechaFinMesSiguiente = mesSiguienteUltimoDia(periodo)
            i = 0
            correlativo = 1
            cantidadCampanas = 0
            complementoCliente = extraerComplementoCliente(len(LOG_PROCESO_PROACTIVA), periodo)
            LOG_PROCESO_PROACTIVA.update(LOG_COMPLEMENTO_CLIENTE)
            ejecutivosExistentesDb = buscarEjecutivosDb2()
            LOG_PROCESO_PROACTIVA.setdefault('INICIO_LECTURA_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: '-----------------------------------------------------' })
            LOG_PROCESO_PROACTIVA.setdefault('ENCABEZADO_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Encabezado del Archivo: %s OK' % archivoEntrada})
            LOG_PROCESO_PROACTIVA.setdefault('INICIO_CELDAS_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Proactiva' , unit=' Fila'):

                i += 1
                if i >= 2:

                    nombreCliente = str(fila[columna['NOMBRE_CLIENTE']].value)
                    nombreCampana = str(fila[columna['NOMBRE_DE_CAMPAÑA']].value)
                    codigoEjecutivo = int(fila[columna['CODIGO_EJECUTIVO']].value)
                    estado = str(fila[columna['ESTADO']].value)
                    estadoRetencion = fila[columna['ESTADO_RETENCION']].value
                    campanaId = str(fila[columna['CAMAPAÑA_ID']].value)
                    estadoUltimaTarea = fila[columna['ESTADO_ULTIMA_TAREA']].value
                    numeroPoliza = formatearNumeroPoliza(fila[columna['NRO_POLIZA']].value)
                    numeroPolizaOriginal = fila[columna['NRO_POLIZA']].value
                    idCliente = formatearIdCliente(nombreCliente)
                    pk = '%s_%s_%s' % (nombreCliente.replace(" ", ""), codigoEjecutivo, numeroPoliza)

                    if numeroPoliza is None:
                        celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_NULO': '%s;Numero de poliza NULL;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_CREACION']])
                    fechaCierre = None
                    fechaExpiracionCoret = None
                    if fila[columna['EXPIRACION_CORET']].value is not None:
                        fechaExpiracionCoret = setearFechaCelda(fila[columna['EXPIRACION_CORET']])
                    if fila[columna['FECHA_CIERRE']].value is not None:
                        fechaCierre = setearFechaCelda(fila[columna['FECHA_CIERRE']])

                    estadoValido = getEstado(fila[columna['ESTADO']])
                    estadoUtValido = getEstadoUt(fila[columna['ESTADO_ULTIMA_TAREA']])

                    if type(fechaCreacion) is not datetime.date:
                        valorErroneo = str(fila[columna['FECHA_CREACION']].value)
                        celdaCoordenada = setearCelda2(fila[0:columna['FECHA_CREACION']+1], len(fila[0:columna['FECHA_CREACION']])-1, i)
                        mensaje = '%s;FECHA_CREACION no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'FECHA_CREACION': mensaje})
                        continue

                    if estado != 'Sin Gestion' and  type(fechaCierre) is not datetime.date:
                        valorErroneo = str(fila[columna['FECHA_CIERRE']].value)
                        celdaCoordenada = setearCelda2(fila[0:columna['FECHA_CIERRE']+1], len(fila[0:columna['FECHA_CIERRE']])-1, i)
                        mensaje = '%s;FECHA_CIERRE no es una fecha valida;%s' % (celdaCoordenada,valorErroneo)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'FECHA_CIERRE': mensaje})
                        continue

                    if estado == 'Sin Gestion' and type(fechaExpiracionCoret) is not datetime.date:
                        valorErroneo = str(fila[columna['EXPIRACION_CORET']].value)
                        celdaCoordenada = setearCelda2(fila[0:columna['EXPIRACION_CORET']+1], len(fila[0:columna['EXPIRACION_CORET']])-1, i)
                        mensaje = '%s;EXPIRACION_CORET no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'FECHA_EXPIRACION': mensaje})
                        continue

                    if type(estadoValido) is not int:
                        valorErroneo = str(fila[columna['ESTADO']].value)
                        celdaCoordenada = setearCelda2(fila[0:columna['ESTADO']+1], len(fila[0:columna['ESTADO']])-1, i)
                        mensaje = '%s;No existe Estado;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'ERROR_ESTADO': mensaje})
                        continue

                    if type(estadoUtValido) is not int:
                        valorErroneo = str(fila[columna['ESTADO_ULTIMA_TAREA']].value)
                        celdaCoordenada = setearCelda2(fila[0:columna['ESTADO_ULTIMA_TAREA']+1], len(fila[0:columna['ESTADO_ULTIMA_TAREA']])-1, i)
                        mensaje = '%s;No existe EstadoUltimaTarea;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'ERROR_ESTADOUT': mensaje})
                        continue

                    if estado != 'Sin Gestion' and fechaCierre >= fechaIncioMes and fechaCierre <= fechaFinMes or estado == 'Sin Gestion' and fechaExpiracionCoret >= fechaIncioMes or estado == 'Sin Gestion' and fechaExpiracionCoret >= fechaIncioMes and fechaExpiracionCoret <= fechaFinMesSiguiente:

                        if ejecutivosExistentesDb.get(codigoEjecutivo):
                            idEmpleado = ejecutivosExistentesDb[codigoEjecutivo]['CODIGO_EMPLEADO']
                        else:
                            celdaCoordenada = setearCelda2(fila[0:columna['CODIGO_EJECUTIVO']+1], len(fila[0:columna['CODIGO_EJECUTIVO']])-1, i)
                            mensaje = '%s;Ejecutivo no existe en la DB;%s' % (celdaCoordenada, codigoEjecutivo)
                            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'RUT_NO_EXISTE': mensaje})
                            continue

                        valoresCampanas = [nombreCliente, fechaCreacion, nombreCampana, numeroPoliza, fechaCierre, estadoRetencion, estadoUltimaTarea]
                        cantidadCampanas += agregarCampanasPorEjecutivo(idEmpleado, valoresCampanas)
                        cobranzaPro = 0
                        pacpatPro = 0
                        if estado == 'Terminado con Exito':
                            if not complementoCliente.get(numeroPoliza):
                                celdaCoordenada = setearCelda2(fila[columna['NRO_POLIZA']], 0)
                                mensaje = '%s;Poliza no existe en ComplmentoCliente;%s_%s' % (celdaCoordenada, numeroPoliza, campanaId)
                                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_NO_EXISTE': mensaje})
                                continue
                            valoresEntrada = {'ESTADO_RETENCION': estadoRetencion, 'NOMBRE_CAMPAÑA': nombreCampana, 'NUMERO_POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CIERRE': fechaCierre, 'ID_EMPLEADO': idEmpleado, 'NUMER_POLIZA_ORIGINAL': numeroPolizaOriginal, 'CAMPAÑA_ID': campanaId, 'ESTADO_VALIDO': estadoValido, 'ESTADO_VALIDOUT': estadoUtValido, 'FECHA_INICIO_MES': fechaIncioMes, 'CELDA_NROPOLIZA': fila[columna['NRO_POLIZA']]}
                            cobranzaPro, pacpatPro = validarRetencionesPolizas(valoresEntrada, complementoCliente)

                        if filaSalidaXls.get(pk):
                            if estado == 'Terminado con Exito' and filaSalidaXls[pk]['ESTADO_PRO'] != 1:
                                filaSalidaXls.pop(pk)
                            elif estado == 'Pendiente' and filaSalidaXls[pk]['ESTADO_PRO'] != 1 and listaEstadoContactado.get(estadoUltimaTarea) or estado == 'Terminado sin Exito' and filaSalidaXls[pk]['ESTADO_PRO'] != 1 and listaEstadoContactado.get(estadoUltimaTarea):
                                filaSalidaXls.pop(pk)
                            elif estado != 'Sin Gestion':
                                if filaSalidaXls[pk]['ESTADO_PRO'] == 0 or filaSalidaXls[pk]['ESTADO_PRO'] != 0 and listaEstadoContactado.get(estadoUltimaTarea):
                                    filaSalidaXls.pop(pk)
                                else:
                                    celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                                    mensaje = '%s;Poliza duplicada;%s_vs_%s' % (celdaCoordenada, estadoValido, filaSalidaXls[pk]['ESTADO_PRO'])
                                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_DUPLICADA': mensaje})
                                    continue
                            else:
                                celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                                mensaje = '%s;Poliza duplicada;%s_vs_%s' % (celdaCoordenada, estadoValido, filaSalidaXls[pk]['ESTADO_PRO'])
                                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_DUPLICADA': mensaje})
                                continue

                        filaSalidaXls[pk] = {'COBRANZA_PRO': cobranzaPro, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': pacpatPro, 'PACPAT_REL_PRO': 0, 'ESTADO_PRO': estadoValido, 'ESTADO_UT_PRO': estadoUtValido, 'ID_EMPLEADO': idEmpleado, 'CAMPAÑA_ID': campanaId, 'CAMPANA': nombreCampana, 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente}
                        correlativo += 1

            if insertarPeriodoCampanaEjecutivos(campanasPorEjecutivos, fechaIncioMes):
                if insertarCampanaEjecutivos(campanasPorEjecutivos, fechaIncioMes):
                    mensaje = 'InsertarCampanaEjecutivos;Se insertaron correctamente: %s Campañas(s)' % (cantidadCampanas)
                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'INSERTAR_CAMPANAS_EJECUTIVOS': mensaje})
                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'INSERTAR_CAMPANAS_EJECUTIVOS': '-----------------------------------------------------' })

            if len(polizasNoAprobadas) > 0:
                insertarPolizasReliquidar = insertarPolizaNoAprobada(polizasNoAprobadas)
                if insertarPolizasReliquidar:
                    mensaje = 'InsertPolizasReliquidar;Se insertaron correctamente: %s Poliza(s)' % (len(polizasNoAprobadas))
                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'INSERTAR_POLIZAS_RELIQUIDAR': mensaje})

            filaSalidaXls.update(polizasReliquidadas(periodo, complementoCliente))

            LOG_PROCESO_PROACTIVA.setdefault('FIN_CELDAS_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivoEntrada, len(tuple(hoja.rows)))})
            LOG_PROCESO_PROACTIVA.setdefault('PROCESO_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Proceso del Archivo: %s Finalizado' % archivoEntrada})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_PROACTIVA.setdefault('ENCABEZADO_PROACTIVA', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivoEntrada, e)
        LOG_PROCESO_PROACTIVA.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_PROACTIVA)+1: errorMsg})
        LOG_PROCESO_PROACTIVA.setdefault('PROCESO_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Error al procesar Archivo: %s' % archivoEntrada})
        return False, False

