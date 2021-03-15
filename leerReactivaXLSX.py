from openpyxl import load_workbook
from tqdm import tqdm
import datetime
from conexio_db import conectorDB

from validaciones_texto import validarEncabezadoXlsx, setearCelda, setearFechaCelda, primerDiaMes, ultimoDiaMes, setearFechaInput, formatearFechaMesAnterior, setearCelda2, formatearIdCliente, formatearRutGion
from diccionariosDB import buscarEjecutivosDb, buscarPolizasReliquidar, buscarPolizasReliquidarAll
from config_xlsx import REACTIVA_CONFIG_XLSX, PATH_XLSX, listaEstadoContactado, listaEstadoNoContactado

from escribir_txt import salidaArchivoTxt, salidaLogTxt

from complementoCliente import extraerComplementoCliente, LOG_COMPLEMENTO_CLIENTE

LOG_PROCESO_REACTIVA = dict()

def extraerBaseCertificacion():
    pathXlsxEntrada = '../test_xls/REACTIVA/'
    # pathXlsxEntrada = PATH_XLSX
    archivoBaseCertificacion = REACTIVA_CONFIG_XLSX['ARCHIVO_BASE_CERTIFICACION']
    archivo = '%s%s.xlsx' % (pathXlsxEntrada, archivoBaseCertificacion['NOMBRE_ARCHIVO'])
    LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INICIO_BASE_CERTIFICACION': 'Iniciando proceso de lectura del Archivo: %s' % archivo})
    try:
        encabezadoXls = archivoBaseCertificacion['ENCABEZADO']
        celda = archivoBaseCertificacion['COLUMNAS']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        baseCertificado = dict()
        i = 0
        validarArchivo = validarEncabezadoXlsx(hoja['A1:X1'], encabezadoXls, archivo)

        if type(validarArchivo) is not dict:
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1 , {'ENCABEZADO_BASE_CERTIFICACION': 'Encabezado del Archivo: %s OK' % archivo})
            for fila in tqdm(iterable=hoja.rows, total= len(tuple(hoja.rows)), desc= 'Leyendo BaseCertificacion' , unit=' Fila'):

                i += 1
                if i >= 2:

                    numeroPoliza = fila[celda['NRO_POLIZA']].value
                    ejecutivo = str(fila[celda['EJECUTIVO']].value)
                    canal = str(fila[celda['CANAL']].value)
                    tipoCertificacion = str(fila[celda['TIPO_CERTIFICACION']].value)
                    fechaLlamado = setearFechaCelda(fila[celda['FECHA_LLAMADO']])

                    if numeroPoliza is None:
                        celdaCoordenada = setearCelda2(fila[celda['NRO_POLIZA']],0)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'POLIZA_NULO': '%s;Numero de poliza NULL BaseCertificado;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    if tipoCertificacion.upper() != 'GRABACIÓN CERTIFICADA':
                        continue

                    if type(fechaLlamado) is not datetime.date:
                        valorErroneo = fila[celda['FECHA_LLAMADO']].value
                        celdaCoordenada = setearCelda2(fila[0:celda['FECHA_LLAMADO']+1], len(fila[0:celda['FECHA_LLAMADO']])-1, i)
                        mensaje = '%s;FECHA_LLAMADO no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FECHA_LLAMADO': mensaje})
                        continue

                    baseCertificado[numeroPoliza] = {'NRO_POLIZA': str(numeroPoliza), 'FECHA_LLAMADO': fechaLlamado, 'EJECUTIVO': ejecutivo, 'CANAL': canal, 'TIPO_CERTIFICACION': tipoCertificacion}

            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LECTURA_BASE_CERTIFICACION': 'Lectura del Archivo: %s Finalizado - %s Filas' % (archivo, len(tuple(hoja.rows)))})
            return baseCertificado
        else:
            LOG_PROCESO_REACTIVA.setdefault('ENCABEZADO_BASE_CERTIFICACION', validarArchivo)
    except Exception as e:
        errorMsg = 'Error al leer archivo;%s | %s' % (archivo, e)
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LECTURA_BASE_CERTIFICACION': errorMsg})
        raise

def validarEstadoReact(estadoRetencion, estado):

    listaEstado = {'Terminado con Exito': 1, 'Pendiente': 2, 'Terminado sin Exito': 3, 'Sin Gestion': 4, 'No gestionable': 5}
    if estadoRetencion == 'Mantiene su producto' or estadoRetencion is None and estado == 'Terminado con Exito':
        estadoValidoReact = listaEstado.get('Terminado con Exito')
    elif estadoRetencion is None:
        estadoValidoReact = listaEstado.get(estado)
    elif estadoRetencion != 'Mantiene su producto':
        estadoValidoReact = listaEstado.get('No gestionable')
    return estadoValidoReact

def validarContactoReact(saliente, estadoRetencion, estado, estadoUt):

    listaEstadoRetencion = REACTIVA_CONFIG_XLSX['ESTADOS_RETENCION']
    if saliente == 0:
        contactoReact = 1
    elif saliente == 1:
        if estadoRetencion == 'Mantiene su producto':
            contactoReact = 1
        elif estadoRetencion is None:
            if estado == 'Sin Gestion':
                contactoReact = 0
        elif estadoRetencion != 'Mantiene su producto':
            if estadoUt is None:
                if listaEstadoRetencion.get(estadoRetencion):
                    contactoReact = 0
                else:
                    if estado == 'Terminado con Exito':
                        contactoReact = 1
                    else:
                        contactoReact = 0
            elif listaEstadoContactado.get(estadoUt):
                contactoReact = 1
            elif listaEstadoNoContactado.get(estadoUt):
                contactoReact = 0
            else:
                mensaje = 'Error;ESTADO_UT no existe;%s' % estadoUt
                LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ERROR_ESTAD-UT': mensaje})
    return contactoReact

def campanaCanal(campana):
    if campana == 'Inbound CoRet':
        valorCamapana = 0
    else:
        valorCamapana = 1
    return valorCamapana
        

def leerArchivoReactiva(archivoEntrada, periodo, fechaInicioEntrada, fechaFinEntrada):

    try:
        archivoSalida = REACTIVA_CONFIG_XLSX['SALIDA_TXT']
        encabezadoXls = REACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX']
        coordenadaEncabezado = REACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO']
        columna = REACTIVA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        complementoCliente = extraerComplementoCliente(len(LOG_PROCESO_REACTIVA))
        LOG_PROCESO_REACTIVA.update(LOG_COMPLEMENTO_CLIENTE)
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'DIVISOR_PROCESO': '-----------------------------------------------------'})
        baseCertificacion = extraerBaseCertificacion()
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'DIVISOR_PROCESO': '-----------------------------------------------------'})
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INICIO_LECTURA_REACTIVA': 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        archivo_correcto = validarEncabezadoXlsx(hoja[coordenadaEncabezado], encabezadoXls, archivoEntrada)

        if type(archivo_correcto) is not dict:
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ENCABEZADO_REACTIVA': 'Encabezado del Archivo: %s OK' % archivoEntrada})
            campanaDescripcion = {0 : 'Inbound CoRet', 1: 'Outbound CoRet'}
            campanaIdDuplicado = dict()
            gestionReactTxt = dict()
            polizaExitoRepetido = dict()
            polizaReactTxt = dict()
            certificacionReactTxt = dict()

            fechaInicioPeriodo = setearFechaInput(fechaInicioEntrada)
            fechaFinPeriodo = setearFechaInput(fechaFinEntrada)
            # periodoMesAnterior = formatearFechaMesAnterior(periodo)
            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            ejecutivosExistentesDb = buscarEjecutivosDb()
            i = 0
            correlativo = 1
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INICIO_CELDAS_REACTIVA': 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Reactiva' , unit=' Fila'):

                i += 1
                if i >= 2:

                    salienteEntrada = fila[columna['LLAMADA_SALIENTE']].value
                    estado = str(fila[columna['ESTADO']].value)
                    estadoUt = fila[columna['ESTADO_ULTIMA_TAREA']].value
                    estadoRetencion = fila[columna['ESTADO_RETENCION']].value
                    numeroPoliza = str(fila[columna['NRO_POLIZA']].value)
                    nombreCliente = str(fila[columna['NOMBRE_CLIENTE']].value)
                    nombreEjecutivo = str(fila[columna['NOMBRE_EJECUTIVO']].value).lower()
                    campanaId = str(fila[columna['CAMAPAÑA_ID']].value)
                    idCliente = formatearIdCliente(nombreCliente)
                    pk = '%s_%s' % (campanaId, numeroPoliza)

                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_CREACION']])
                    fechaCierre = setearFechaCelda(fila[columna['FECHA_CIERRE']])

                    if salienteEntrada is None:
                        celdaCoordenada = setearCelda2(fila[0:columna['LLAMADA_SALIENTE']+1], len(fila[0:columna['LLAMADA_SALIENTE']])-1, i)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LLAMADA_SALIENTE': '%s;Campaña InBound/OutBound NULL;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    if numeroPoliza is None:
                        celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'POLIZA_NULO': '%s;Numero de poliza NULL;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    if type(fechaCreacion) is not datetime.date:
                        valorErroneo = fila[columna['FECHA_CREACION']].value
                        celdaCoordenada = setearCelda2(fila[0:columna['FECHA_CREACION']+1], len(fila[0:columna['FECHA_CREACION']])-1, i)
                        mensaje = '%s;FECHA_CREACION no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FECHA_CREACION': mensaje})
                        continue

                    if type(fechaCierre) is not datetime.date:
                        fechaCierre = None

                    if not ejecutivosExistentesDb.get(nombreEjecutivo):
                        valorErroneo = fila[columna['NOMBRE_EJECUTIVO']].value
                        celdaCoordenada = setearCelda2(fila[0:columna['NOMBRE_EJECUTIVO']+1], len(fila[0:columna['NOMBRE_EJECUTIVO']])-1, i)
                        mensaje = '%s;NOMBRE_EJECUTIVO no existe en la DB;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'NOMBRE_EJECUTIVO': mensaje})
                        continue

                    saliente = int(salienteEntrada)
                    if saliente == 0 and fechaCreacion >= fechaIncioMes and fechaCreacion <= fechaFinMes or saliente == 1 and fechaCreacion >= fechaInicioPeriodo and fechaCreacion <= fechaFinPeriodo:

                        estadoValidoReact = validarEstadoReact(estadoRetencion, estado)
                        contactoReact = validarContactoReact(saliente, estadoRetencion, estado, estadoUt)
                        rut = ejecutivosExistentesDb.get(nombreEjecutivo)
                        rutFormateado = formatearRutGion(rut['RUT'])

                        if campanaIdDuplicado.get(campanaId):
                            pkDataGestion = campanaIdDuplicado[campanaId]['PK']
                            if contactoReact == 1 and gestionReactTxt[pkDataGestion]['CONTACTO_REACT'] == 0:
                                datosActualizados = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 1, 'RUT': rutFormateado, 'ID_CAMPANA': campanaId, 'CAMPANA': campanaDescripcion.get(saliente), 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                                gestionReactTxt[pkDataGestion].update(datosActualizados)
                        else:
                            campanaIdDuplicado[campanaId] = {'PK': pk}

                        if not gestionReactTxt.get(pk):
                            gestionReactTxt[pk] = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 1, 'RUT': rutFormateado, 'ID_CAMPANA': campanaId, 'CAMPANA': campanaDescripcion.get(saliente), 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                        else:
                            if estadoValidoReact == 1 and gestionReactTxt[pk]['ESTADO_VALIDO_REACT'] != 1:
                                datosActualizados = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 1, 'RUT': rutFormateado, 'ID_CAMPANA': campanaId, 'CAMPANA': campanaDescripcion.get(saliente), 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                                gestionReactTxt[pk].update(datosActualizados)
                            elif contactoReact == 1 and gestionReactTxt[pk]['CONTACTO_REACT'] == 0:
                                datosActualizados = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 1, 'RUT': rutFormateado, 'ID_CAMPANA': campanaId, 'CAMPANA': campanaDescripcion.get(saliente), 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                                gestionReactTxt[pk].update(datosActualizados)
                            elif estadoValidoReact != 4:
                                datosActualizados = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 1, 'RUT': rutFormateado, 'ID_CAMPANA': campanaId, 'CAMPANA': campanaDescripcion.get(saliente), 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                                gestionReactTxt[pk].update(datosActualizados)
                        
                        if polizaExitoRepetido.get(numeroPoliza):
                            pkDataGestion = polizaExitoRepetido[numeroPoliza]['PK']
                            if saliente == 0 and fechaCreacion > gestionReactTxt[pkDataGestion]['FECHA_CREACION']:
                                gestionReactTxt[pkDataGestion]['EXITO_REPETIDO_REACT'] = 0
                            elif saliente == 1:
                                if type(fechaCierre) is datetime.date and type(gestionReactTxt[pkDataGestion]['FECHA_CIERRE']) is datetime.date:
                                    if fechaCierre > gestionReactTxt[pkDataGestion]['FECHA_CIERRE']:
                                        gestionReactTxt[pkDataGestion]['EXITO_REPETIDO_REACT'] = 0
                                else:
                                    valorErroneo =  '%s-VS-%s' % (fechaCierre, gestionReactTxt[pkDataGestion]['FECHA_CIERRE'])
                                    celdaCoordenada = setearCelda2(fila[0:columna['FECHA_CIERRE']+1], len(fila[0:columna['FECHA_CIERRE']])-1, i)
                                    mensaje = '%s;FECHA_CIERRE no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                                    LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FECHA_CIERRE': mensaje})
                        else:
                            if estadoValidoReact == 1:
                                polizaExitoRepetido[numeroPoliza] = {'PK': pk}

                        if estadoValidoReact == 1:
                            if complementoCliente.get(numeroPoliza) and complementoCliente[numeroPoliza]['ESTADO_POLIZA'] == 'Vigente':
                                polizaReactTxt[numeroPoliza] = {'ESTADO_POLIZA_REACT': 1, 'NUMERO_POLIZA': numeroPoliza}

                        if estadoValidoReact == 1:
                            if baseCertificacion.get(numeroPoliza):

                                nombreBaseCertificacion = str(baseCertificacion[numeroPoliza]['EJECUTIVO']).lower()
                                if not ejecutivosExistentesDb.get(nombreBaseCertificacion):
                                    valorErroneo = baseCertificacion[numeroPoliza]['EJECUTIVO']
                                    mensaje = '%s;EJECUTIVO_BASE_CERIFICACION no existe en la DB;%s' % (celdaCoordenada, valorErroneo)
                                    LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'NOMBRE_EJECUTIVO': mensaje})
                                    continue

                                grabCertificadaReact = 0
                                rutBaseCerfiticacion = ejecutivosExistentesDb.get(nombreBaseCertificacion)
                                campanaBaseCertificacion = campanaCanal(baseCertificacion[numeroPoliza]['CANAL'])
                                fechaLlamado = baseCertificacion[numeroPoliza]['FECHA_LLAMADO']

                                if campanaBaseCertificacion == 0 and fechaLlamado >= fechaIncioMes and fechaLlamado <= fechaFinMes or campanaBaseCertificacion == 1 and fechaLlamado >= fechaInicioPeriodo and fechaLlamado <= fechaFinPeriodo:
                                    grabCertificadaReact = 1

                                certificacionReactTxt[numeroPoliza] = {'GRAB_CERTIFICADA_REACT': grabCertificadaReact, 'RUT': rutBaseCerfiticacion, 'CANPANA': campanaBaseCertificacion, 'POLIZA': numeroPoliza}

                    correlativo += 1

            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FIN_CELDAS_REACTIVA': 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivoEntrada, len(tuple(hoja.rows)))})
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'PROCESO_REACTIVA': 'Proceso del Archivo: %s Finalizado' % archivoEntrada})
            dataSalida = [
                {'NOMBRE_ARCHIVO': archivoSalida['GESTION']['NOMBRE_SALIDA'], 'DATA': gestionReactTxt, 'ENCABEZADO': archivoSalida['GESTION']['ENCABEZADO']},
                {'NOMBRE_ARCHIVO': archivoSalida['POLIZA']['NOMBRE_SALIDA'], 'DATA': polizaReactTxt, 'ENCABEZADO': archivoSalida['POLIZA']['ENCABEZADO']},
                {'NOMBRE_ARCHIVO': archivoSalida['CERTIFICACION']['NOMBRE_SALIDA'], 'DATA': certificacionReactTxt, 'ENCABEZADO': archivoSalida['CERTIFICACION']['ENCABEZADO']}
            ]
            return dataSalida
        else:
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ENCABEZADO_REACTIVA': archivo_correcto})
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivoEntrada, e)
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LECTURA_ARCHIVO': errorMsg})
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'PROCESO_REACTIVA': 'Error al procesar Archivo: %s' % archivoEntrada})
        return False, False
        # raise

x = leerArchivoReactiva('../test_xls/REACTIVA/Gestion Reactiva.xlsx', '202006', '20200615', '202007030')
print(x[2])
# print(LOG_PROCESO_REACTIVA)
salidaLogTxt('../test_xls/REACTIVA/reactiva.log', LOG_PROCESO_REACTIVA)
# print(salidaArchivoTxt('../test_xls/REACTIVA/%s%s.txt' % (REACTIVA_CONFIG_XLSX['SALIDA_TXT'],'202010'), x, y))
