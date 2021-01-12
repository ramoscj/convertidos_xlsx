from openpyxl import load_workbook
from tqdm import tqdm
import datetime
from conexio_db import conectorDB

from validaciones_texto import validarEncabezadoXlsx, setearCelda, setearFechaCelda, primerDiaMes, ultimoDiaMes, mesSiguienteUltimoDia, setearCelda2, formatearFechaMesAnterior, formatearNumeroPoliza
from diccionariosDB import buscarEjecutivosDb, buscarPolizasReliquidar, buscarPolizasReliquidarAll
from config_xlsx import PROACTIVA_CONFIG_XLSX, PATH_XLSX

from escribir_txt import salidaArchivoTxt, salidaLogTxt

LOG_PROCESO_PROACTIVA = dict()

def extraerComplementoCliente():
    pathXlsxEntrada = 'test_xls/'
    # pathXlsxEntrada = PATH_XLSX
    archivo = '%s%s.xlsx' % (pathXlsxEntrada, PROACTIVA_CONFIG_XLSX['ENTRADA_COMPLEMENTO_CLIENTE'])
    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'INICIO_LECTURA_PROPIETARIOS': 'Iniciando proceso de lectura del Archivo: %s' % archivo})
    try:
        encabezadoXls = PROACTIVA_CONFIG_XLSX['ENCABEZADO_COMPLEMENTO_CLIENTE']
        celda = PROACTIVA_CONFIG_XLSX['COLUMNAS_COMPLEMENTO_CLIENTE']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        complementoCliente = dict()
        validarArchivo = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivo)

        if type(validarArchivo) is not dict:
            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1 , {'ENCABEZADO_COMPLEMENTOCLIENTE': 'Encabezado del Archivo: %s OK' % archivo})
            for fila in tqdm(iterable=hoja.rows, total= len(tuple(hoja.rows)), desc= 'Leyendo ComplementoCliente' , unit=' Fila'):

                nroPoliza = str(fila[celda['NRO_POLIZA']].value)
                fecUltPago = None
                fecMandato = None
                if fila[celda['FEC_ULT_PAG']].value is not None:
                    fecUltPago = setearFechaCelda(fila[celda['FEC_ULT_PAG']])
                if fila[celda['FECHA_MANDATO']].value is not None:
                    fecMandato = setearFechaCelda(fila[celda['FECHA_MANDATO']])
                complementoCliente[nroPoliza] = {'NRO_CERT': str(fila[celda['NRO_CERT']].value), 'FEC_ULT_PAG': fecUltPago, 'ESTADO_MANDATO': fila[celda['ESTADO_MANDATO']].value, 'FECHA_MANDATO': fecMandato}

            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'LECTURA_COMPLEMENTOCLIENTE': 'Lectura del Archivo: %s Finalizado' % archivo})
            return complementoCliente
        else:
            LOG_PROCESO_PROACTIVA.setdefault('ENCABEZADO_COMPLEMENTOCLIENTE', validarArchivo)
    except Exception as e:
        errorMsg = 'Error al leer archivo;%s | %s' % (archivo, e)
        LOG_PROCESO_PROACTIVA.setdefault('LECTURA_COMPLEMENTOCLIENTE' , errorMsg)
        raise

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
        sql = """INSERT INTO retenciones_por_reliquidar (rut_ejecutivo, id_cliente, numero_poliza, campana_id, cobranza_pro, cobranza_rel_pro, pacpat_pro, pacpat_rel_pro, estado_pro, estado_ut_pro, fecha_proceso, fecha_reliquidacion, fecha_cierre) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?);"""
        cursor.executemany(sql, dataPolizas)
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
        if poliza['COBRANZA_PRO'] > 0:
            numeroPolizaCertificado = estadoCertificadoPoliza(poliza['POLIZA'])
            numeroPoliza = formatearNumeroPoliza(poliza['POLIZA'])
            if complementoCliente.get(numeroPoliza):
                cobranzaPro = aprobarCobranza(numeroPolizaCertificado, poliza['FECHA_CIERRE'], complementoCliente[numeroPoliza]['NRO_CERT'], complementoCliente[numeroPoliza]['FEC_ULT_PAG'])
                if cobranzaPro == 1:
                    polizasAprobadaReliquidar[numeroPoliza] = {'CRR': correlativo, 'COBRANZA_PRO': poliza['COBRANZA_PRO'], 'COBRANZA_REL_PRO': poliza['COBRANZA_REL_PRO'], 'PACPAT_PRO': poliza['PACPAT_PRO'], 'PACPAT_REL_PRO': poliza['PACPAT_REL_PRO'], 'ESTADO_PRO': poliza['ESTADO_PRO'], 'ESTADO_UT_PRO': poliza['ESTADO_UT_PRO'], 'RUT': poliza['RUT'], 'CAMPAÑA_ID': poliza['CAMPAÑA_ID'], 'POLIZA': numeroPoliza, 'ID_CLIENTE': poliza['ID_CLIENTE']}
                    correlativo += 1
                else:
                    mensaje = 'PolizaReliquidacion;No cumple condicion de retencion COBRO para Reliquidacion;%s' % (numeroPoliza)
                    LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA_RELIQUIDACION': mensaje})
    print(polizasAprobadaReliquidar)

    # if pacpatPro > 0:
    #     estadoMandato = complementoCliente[numeroPoliza]['ESTADO_MANDATO']
    #     fecMandato = complementoCliente[numeroPoliza]['FECHA_MANDATO']
    #     if estadoMandato is not None:
    #         pacpatPro = aprobarActivacion(estadoMandato, fecMandato, fechaCierre)
    #     else:
    #         pacpatPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
    #     if pacpatPro == 0:
    #         if not polizasReliquidarDb.get(pk2):
    #             polizasNoAprobadas.append((rut, idCiente, numeroPolizaOriginal, campanaId, 1, 0, pacpatPro, 0, estadoValido, estadoUtValido, fechaIncioMes, fechaCierre))
    #         else:
    #             celdaCoordenada = setearCelda2(fila[columna['NRO_POLIZA']], 0)
    #             mensaje = '%s;Poliza no cumple condicion de retencion ACTIVACION y esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
    #             LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA': mensaje})
    #         continue

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
            polizasNoAprobadas = []

            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            fechaFinMesSiguiente = mesSiguienteUltimoDia(periodo)
            i = 0
            correlativo = 1
            complementoCliente = extraerComplementoCliente()
            ejecutivosExistentesDb = buscarEjecutivosDb()
            polizasReliquidarDb = buscarPolizasReliquidarAll()
            LOG_PROCESO_PROACTIVA.setdefault('INICIO_CELDAS_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            filaSalidaXls = dict()
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
                    estadoUltimaTarea = fila[columna['ESTADO_ULTIMA_TAREA']]
                    numeroPoliza = formatearNumeroPoliza(fila[columna['NRO_POLIZA']].value)
                    numeroPolizaOriginal = fila[columna['NRO_POLIZA']].value
                    pk = '%s_%s_%s' % (nombreCliente, nombreEjecutivo, numeroPoliza)

                    if numeroPoliza is None:
                        celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_NULO': '%s;Numero de poliza NULL;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    if ejecutivosExistentesDb.get(nombreEjecutivo):
                        rut = rut = ejecutivosExistentesDb[nombreEjecutivo]['RUT']
                    else:
                        celdaCoordenada = setearCelda2(fila[0:columna['NOMBRE_EJECUTIVO']+1], len(fila[0:columna['NOMBRE_EJECUTIVO']])-1, i)
                        mensaje = '%s;Ejecutivo no existe en la DB;%s' % (celdaCoordenada, nombreEjecutivo)
                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'RUT_NO_EXISTE': mensaje})
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

                        cobranzaPro = 0
                        pacpatPro = 0
                        if not estado == 'Terminado con Exito':
                            continue

                        listaConsiderarRetencion = {'Mantiene su producto': 1, 'Realiza pago en línea': 2, 'Realiza Activación PAC/PAT': 3}
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

                        if complementoCliente.get(numeroPoliza):

                            numeroPolizaCertificado = estadoCertificadoPoliza(numeroPoliza)
                            fecUltimoPago = complementoCliente[numeroPoliza]['FEC_ULT_PAG']
                            numeroPolizaCliente = complementoCliente[numeroPoliza]['NRO_CERT']
                            idCiente = 'xx'
                            pk2 = '%s_%s_%s' % (idCiente, rut, numeroPolizaOriginal)
                            if cobranzaPro > 0:
                                cobranzaPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
                                if cobranzaPro == 0:
                                    if not polizasReliquidarDb.get(pk2):
                                        print(campanaId)
                                        polizasNoAprobadas.append((rut, idCiente, numeroPolizaOriginal, campanaId, 1, 0, pacpatPro, 0, estadoValido, estadoUtValido, fechaIncioMes, fechaCierre))
                                    else:
                                        celdaCoordenada = setearCelda2(fila[columna['NRO_POLIZA']], 0)
                                        mensaje = '%s;Poliza no cumple condicion de retencion COBRO y esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
                                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA': mensaje})
                                    continue

                            if pacpatPro > 0:
                                estadoMandato = complementoCliente[numeroPoliza]['ESTADO_MANDATO']
                                fecMandato = complementoCliente[numeroPoliza]['FECHA_MANDATO']
                                if estadoMandato is not None:
                                    pacpatPro = aprobarActivacion(estadoMandato, fecMandato, fechaCierre)
                                else:
                                    pacpatPro = aprobarCobranza(numeroPolizaCertificado, fechaCierre, numeroPolizaCliente, fecUltimoPago)
                                if pacpatPro == 0:
                                    if not polizasReliquidarDb.get(pk2):
                                        polizasNoAprobadas.append((rut, idCiente, numeroPolizaOriginal, campanaId, 1, 0, pacpatPro, 0, estadoValido, estadoUtValido, fechaIncioMes, fechaCierre))
                                    else:
                                        celdaCoordenada = setearCelda2(fila[columna['NRO_POLIZA']], 0)
                                        mensaje = '%s;Poliza no cumple condicion de retencion ACTIVACION y esta duplicada en la DB;%s' % (celdaCoordenada, numeroPoliza)
                                        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'PROCESO_COBRANZA': mensaje})
                                    continue

                        else:
                            celdaCoordenada = setearCelda2(fila[columna['NRO_POLIZA']], 0)
                            mensaje = '%s;Poliza no existe en ComplmentoCliente;%s' % (celdaCoordenada, numeroPoliza)
                            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1, {'POLIZA_NO_EXISTE': mensaje})
                            continue


                        # PREGUNTA PENDIENTE
                        if filaSalidaXls.get(pk):
                            print(campanaId)
                            continue
                        filaSalidaXls[pk] = {'CRR': correlativo, 'COBRANZA_PRO': cobranzaPro, 'COBRANZA_REL_PRO': 0, 'PACPAT_PRO': pacpatPro, 'PACPAT_REL_PRO': 0, 'ESTADO_PRO': estadoValido, 'ESTADO_UT_PRO': estadoUtValido, 'RUT': rut, 'CAMPAÑA_ID': campanaId, 'POLIZA': numeroPoliza, 'ID_CLIENTE': 'xx'}
                        correlativo += 1

            if len(polizasNoAprobadas) > 0:
                print(insertarPolizaNoAprobada(polizasNoAprobadas))
            mesAnterior = formatearFechaMesAnterior(periodo)
            polizasReliquidadas(mesAnterior, complementoCliente, correlativo)

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
        # return False, False
        raise

x, y = leerArchivoProactiva('test_xls/Gestión CoRet Proactiva_2.xlsx', '202009')
salidaLogTxt('test_xls/proactiva.log', LOG_PROCESO_PROACTIVA)
print(salidaArchivoTxt('test_xls/PROACTIVA.txt', x, y))

# print(formatearNumeroPoliza(''))
# print(x)
# print(y)
# print(LOG_PROCESO_PROACTIVA)
