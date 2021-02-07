from openpyxl import load_workbook
from tqdm import tqdm
import datetime
from conexio_db import conectorDB

from validaciones_texto import validarEncabezadoXlsx, setearCelda, setearFechaCelda, primerDiaMes, ultimoDiaMes, mesSiguienteUltimoDia, setearCelda2, formatearFechaMesAnterior, formatearNumeroPoliza, formatearIdCliente, convertirDiccionario
from diccionariosDB import buscarEjecutivosDb, buscarPolizasReliquidar, buscarPolizasReliquidarAll
from config_xlsx import PROACTIVA_CONFIG_XLSX, PATH_XLSX, listaEstadoContactado

from escribir_txt import salidaArchivoTxt, salidaLogTxt, salidaArchivoTxtProactiva

from complementoCliente import extraerComplementoCliente, LOG_COMPLEMENTO_CLIENTE

LOG_PROCESO_PROACTIVA = dict()
polizasNoAprobadas = dict()

def getEstado(celdaEstado):
    listaContactado = {'Pendiente':1 , 'Terminado con Exito': 2 , 'Terminado sin Exito': 3}
    if listaContactado.get(str(celdaEstado.value)):
        return listaContactado[celdaEstado.value]
    elif str(celdaEstado.value) == 'Sin Gestion':
        return 0
    else:
        return False

def getEstadoUt(celdaEstadoUt):
    listaEstadoUt = PROACTIVA_CONFIG_XLSX['LISTA_ULTIMA_TAREA']
    if listaEstadoUt.get(celdaEstadoUt.value):
        return listaEstadoUt[celdaEstadoUt.value]
    elif celdaEstadoUt.value is None:
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

def insertarPolizaNoAprobada(dataPolizas:[]):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasInsertar = convertirDiccionario(dataPolizas)
        sql = """INSERT INTO retenciones_por_reliquidar (rut_ejecutivo, id_cliente, numero_poliza, campana_id, cobranza_pro, cobranza_rel_pro, pacpat_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?);"""
        cursor.executemany(sql, polizasInsertar)
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al insertar polizas para reliquidar: %s - %s' % (e ,e))
    finally:
        cursor.close()
        db.close()

def polizasReliquidadas(mesAnterior, complementoCliente, correlativo):
    polizasParaReliquidar = buscarPolizasReliquidar(mesAnterior)
    polizasAprobadaReliquidar = dict()

    for poliza in polizasParaReliquidar.values():
        numeroPolizaCertificado = estadoCertificadoPoliza(poliza['POLIZA'])
        numeroPoliza = formatearNumeroPoliza(poliza['POLIZA'])
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
            polizasAprobadaReliquidar[numeroPoliza] = {'COBRANZA_PRO': 0, 'COBRANZA_REL_PRO': cobranzaRelPro, 'PACPAT_PRO': 0, 'PACPAT_REL_PRO': pacpatRelPro, 'ESTADO_PRO': poliza['ESTADO_PRO'], 'ESTADO_UT_PRO': poliza['ESTADO_UT_PRO'], 'RUT': poliza['RUT'], 'CAMPAÑA_ID': poliza['CAMPAÑA_ID'], 'POLIZA': numeroPoliza, 'ID_CLIENTE': poliza['ID_CLIENTE']}
            correlativo += 1

    return polizasAprobadaReliquidar

def validarRetencionesPolizas(valoresEntrada: dict, complementoCliente: dict):

    estadoRetencion = valoresEntrada['ESTADO_RETENCION']
    nombreCampana = valoresEntrada['NOMBRE_CAMPAÑA']
    numeroPoliza = valoresEntrada['NUMERO_POLIZA']
    rut = valoresEntrada['RUT']
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
    pk2 = '%s_%s_%s' % (idCiente, rut, numeroPolizaOriginal)
    if cobranzaPro > 0:
        cobranzaPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
        if cobranzaPro == 0:
            if not polizasReliquidarDb.get(pk2):
                if not polizasNoAprobadas.get(pk2):
                    polizasNoAprobadas[pk2] = {'RUT': rut, 'ID_CLIENTE': idCiente, 'NRO_POLIZA': numeroPolizaOriginal, 'ID_CAMPAÑA': campanaId, 'COBRANZA_PRO': 1, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': pacpatPro, 'PACPAT_REL_PRO': 0, 'ESTADO_VALIDO': estadoValido, 'ESTADO_VALIDOUT': estadoUtValido, 'FECHA_INICIO_MES': fechaIncioMes, 'FECHA_CIERRE': fechaCierre}
            else:
                celdaCoordenada = setearCelda2(celdaNroPoliza, 0)
                mensaje = '%s;Poliza no cumple condicion de retencion COBRO y esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
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
                    polizasNoAprobadas[pk2] = {'RUT': rut, 'ID_CLIENTE': idCiente, 'NRO_POLIZA': numeroPolizaOriginal, 'ID_CAMPAÑA': campanaId, 'COBRANZA_PRO': cobranzaPro, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': 1, 'PACPAT_REL_PRO': 0, 'ESTADO_VALIDO': estadoValido, 'ESTADO_VALIDOUT': estadoUtValido, 'FECHA_INICIO_MES': fechaIncioMes, 'FECHA_CIERRE': fechaCierre}
                else:
                    polizasNoAprobadas[pk2]['PACPAT_PRO'] = 1
            else:
                celdaCoordenada = setearCelda2(celdaNroPoliza, 0)
                mensaje = '%s;Poliza no cumple condicion de retencion ACTIVACION y esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_ACTIVACION': mensaje})

    return cobranzaPro, pacpatPro

def leerArchivoProactiva(archivoEntrada, periodo):
    try:
        LOG_PROCESO_PROACTIVA.setdefault('INICIO_LECTURA_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        encabezadoXls = PROACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = PROACTIVA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = PROACTIVA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        archivo_correcto = validarEncabezadoXlsx(hoja['A1:K1'], encabezadoXls, archivoEntrada)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_PROACTIVA.setdefault('ENCABEZADO_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Encabezado del Archivo: %s OK' % archivoEntrada})
            filaSalidaXls = dict()

            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            fechaFinMesSiguiente = mesSiguienteUltimoDia(periodo)
            i = 0
            correlativo = 1
            complementoCliente = extraerComplementoCliente(len(LOG_PROCESO_PROACTIVA))
            LOG_PROCESO_PROACTIVA.update(LOG_COMPLEMENTO_CLIENTE)
            ejecutivosExistentesDb = buscarEjecutivosDb()
            LOG_PROCESO_PROACTIVA.setdefault('INICIO_CELDAS_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Proactiva' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):

                i += 1
                if i >= 2:

                    nombreCliente = str(fila[columna['NOMBRE_CLIENTE']].value)
                    nombreCampana = str(fila[columna['NOMBRE_DE_CAMPAÑA']].value)
                    nombreEjecutivo = str(fila[columna['NOMBRE_EJECUTIVO']].value).lower()
                    estado = str(fila[columna['ESTADO']].value)
                    estadoRetencion = str(fila[columna['ESTADO_RETENCION']].value)
                    campanaId = str(fila[columna['CAMAPAÑA_ID']].value)
                    estadoUltimaTarea = str(fila[columna['ESTADO_ULTIMA_TAREA']].value)
                    numeroPoliza = formatearNumeroPoliza(fila[columna['NRO_POLIZA']].value)
                    numeroPolizaOriginal = fila[columna['NRO_POLIZA']].value
                    idCliente = formatearIdCliente(nombreCliente)
                    pk = '%s_%s_%s' % (nombreCliente, nombreEjecutivo, numeroPoliza)

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

                        if ejecutivosExistentesDb.get(nombreEjecutivo):
                            rut = ejecutivosExistentesDb[nombreEjecutivo]['RUT']
                        else:
                            celdaCoordenada = setearCelda2(fila[0:columna['NOMBRE_EJECUTIVO']+1], len(fila[0:columna['NOMBRE_EJECUTIVO']])-1, i)
                            mensaje = '%s;Ejecutivo no existe en la DB;%s' % (celdaCoordenada, nombreEjecutivo)
                            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'RUT_NO_EXISTE': mensaje})
                            continue

                        cobranzaPro = 0
                        pacpatPro = 0
                        if estado == 'Terminado con Exito':
                            if not complementoCliente.get(numeroPoliza):
                                celdaCoordenada = setearCelda2(fila[columna['NRO_POLIZA']], 0)
                                mensaje = '%s;Poliza no existe en ComplmentoCliente;%s' % (celdaCoordenada, numeroPoliza)
                                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_NO_EXISTE': mensaje})
                                continue
                            valoresEntrada = {'ESTADO_RETENCION': estadoRetencion, 'NOMBRE_CAMPAÑA': nombreCampana, 'NUMERO_POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CIERRE': fechaCierre, 'RUT': rut, 'NUMER_POLIZA_ORIGINAL': numeroPolizaOriginal, 'CAMPAÑA_ID': campanaId, 'ESTADO_VALIDO': estadoValido, 'ESTADO_VALIDOUT': estadoUtValido, 'FECHA_INICIO_MES': fechaIncioMes, 'CELDA_NROPOLIZA': fila[columna['NRO_POLIZA']]}
                            cobranzaPro, pacpatPro = validarRetencionesPolizas(valoresEntrada, complementoCliente)

                        if filaSalidaXls.get(pk):
                            if estado == 'Terminado con Exito':
                                filaSalidaXls.pop(pk)
                            elif estado == 'Pendiente' or estado == 'Terminado sin exito':
                                if listaEstadoContactado.get(estadoUltimaTarea):
                                    filaSalidaXls.pop(pk)
                            elif estado != 'Sin Gestion':
                                filaSalidaXls.pop(pk)
                            else:
                                valorErroneo = pk
                                celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                                mensaje = '%s;Poliza duplicada;%s' % (celdaCoordenada, valorErroneo)
                                LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_DUPLICADA': mensaje})
                                continue
                        filaSalidaXls[pk] = {'COBRANZA_PRO': cobranzaPro, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': pacpatPro, 'PACPAT_REL_PRO': 0, 'ESTADO_PRO': estadoValido, 'ESTADO_UT_PRO': estadoUtValido, 'RUT': rut, 'CAMPAÑA_ID': campanaId, 'CAMPANA': nombreCampana, 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente}
                        correlativo += 1

            if len(polizasNoAprobadas) > 0:
                insertarPolizasReliquidar = insertarPolizaNoAprobada(polizasNoAprobadas)
                if insertarPolizasReliquidar:
                    mensaje = 'InsertPolizasReliquidar;Se insertaron correctamente: %s Poliza(s)' % (len(polizasNoAprobadas))
                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'INSERTAR_POLIZAS_RELIQUIDAR': mensaje})
            mesAnterior = formatearFechaMesAnterior(periodo)
            filaSalidaXls.update(polizasReliquidadas(mesAnterior, complementoCliente, correlativo))

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
        # raise

x, y = leerArchivoProactiva('test_xls/Gestión CoRet Proactiva.xlsx', '202010')
salidaLogTxt('test_xls/proactiva.log', LOG_PROCESO_PROACTIVA)
print(salidaArchivoTxtProactiva('test_xls/%s%s.txt' % (PROACTIVA_CONFIG_XLSX['SALIDA_TXT'],'202010'), x, y))
