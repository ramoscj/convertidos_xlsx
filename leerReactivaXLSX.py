import datetime

from openpyxl import load_workbook
from tqdm import tqdm

from complementoCliente import (LOG_COMPLEMENTO_CLIENTE,
                                extraerComplementoCliente)
from conexio_db import conectorDB
from config_xlsx import PATH_XLSX, REACTIVA_CONFIG_XLSX
from diccionariosDB import (buscarRutEjecutivosDb,
                            listaEstadoRetencionReactiva,
                            listaEstadoUtContacto, listaEstadoUtNoContacto,
                            periodoCampanasReactiva, CampanasPorPeriodoReactiva)
from escribir_txt import (salidaArchivoTxt, salidaArchivoTxtProactiva,
                          salidaLogTxt)
from validaciones_texto import (convertirDataReact,
                                convertirListaReactiva,
                                formatearFechaMesAnterior, formatearIdCliente,
                                formatearNumeroPoliza, formatearRutGion,
                                primerDiaMes, setearCampanasReactiva,
                                setearCelda, setearCelda2, setearFechaCelda,
                                setearFechaInput, ultimoDiaMes,
                                validarEncabezadoXlsx)

LOG_PROCESO_REACTIVA = dict()
campanasPorEjecutivos = dict()
inbound = REACTIVA_CONFIG_XLSX['INBOUND_VALOR']
outbound = REACTIVA_CONFIG_XLSX['OUTBOUND_VALOR']
listaEstadoRetencion = listaEstadoRetencionReactiva()

def extraerBaseCertificacion(archivoCertificacionXls):
    archivo = archivoCertificacionXls
    archivoBaseCertificacion = REACTIVA_CONFIG_XLSX['ARCHIVO_BASE_CERTIFICACION']
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
                    idEmpleado = fila[celda['ID_EMPLEADO']].value
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
                    
                    pk = '{0}_{1}'.format(str(numeroPoliza), str(idEmpleado))
                    baseCertificado[pk] = {'NRO_POLIZA': int(numeroPoliza), 'FECHA_LLAMADO': fechaLlamado, 'ID_EMPLEADO': str(idEmpleado), 'CANAL': canal, 'TIPO_CERTIFICACION': tipoCertificacion}

            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LECTURA_BASE_CERTIFICACION': 'Lectura del Archivo: %s Finalizado - %s Filas' % (archivo, len(tuple(hoja.rows)))})
            return baseCertificado
        else:
            LOG_PROCESO_REACTIVA.setdefault('ENCABEZADO_BASE_CERTIFICACION', validarArchivo)
    except Exception as e:
        errorMsg = 'Error al leer archivo;%s | %s' % (archivo, e)
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LECTURA_BASE_CERTIFICACION': errorMsg})
        raise

def validarEstadoReact(estadoRetencion, estado):
    if estadoRetencion == 'Mantiene su producto' or estadoRetencion is None and estado == 'Terminado con Exito':
        estadoValidoReact = listaEstadoRetencion.get('Terminado con Exito')
    elif estadoRetencion is None:
        estadoValidoReact = listaEstadoRetencion.get(estado)
    elif estadoRetencion != 'Mantiene su producto':
        estadoValidoReact = listaEstadoRetencion.get(estadoRetencion)
    return estadoValidoReact

def validarContactoReact(saliente, estadoRetencion, estado, estadoUt):

    listaEstadoNoContactado = listaEstadoUtNoContacto()
    listaEstadoContactado = listaEstadoUtContacto()
    if saliente == inbound:
        contactoReact = 1
    elif saliente == outbound:
        if estadoRetencion == 'Mantiene su producto':
            contactoReact = 1
        elif estadoRetencion is None:
            if estado == 'Terminado con Exito':
                contactoReact = 1
            else:
                contactoReact = 0
        elif estadoRetencion != 'Mantiene su producto':
            if estadoUt is None:
                if estado == 'Sin Gestion':
                    contactoReact = 1
                else:
                    contactoReact = 0
            elif listaEstadoContactado.get(estadoUt):
                contactoReact = 1
            elif listaEstadoNoContactado.get(estadoUt):
                contactoReact = 0
            else:
                mensaje = 'Error ESTADO_UT;ESTADO_UT no existe;%s' % estadoUt
                LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ERROR_ESTADOUT': mensaje})
    return contactoReact

def campanaCanal(campana):
    if str(campana).upper() == 'INBOUND':
        valorCamapana = 1
    else:
        valorCamapana = 0
    return valorCamapana

def exitoRepetidoPk(numeroPoliza, polizaExitoRepetido, gestionReactTxt):
    pkSalida = 0
    for pkData in polizaExitoRepetido[numeroPoliza]:
        if gestionReactTxt[pkData]['EXITO_REPETIDO_REACT'] == 1:
            pkSalida = pkData
    return pkSalida

def aprobarActualizarRegistro(estado, estadoValidoReact, contactoReact, estadoValidoReactData, contactoReactData):
    controlCambioPk = False
    indiceCambio = 0
    if estadoValidoReact == 1 and estadoValidoReactData != 1:
        controlCambioPk = True
        indiceCambio = 1
    elif contactoReact == 1 and contactoReactData == 0:
        controlCambioPk = True
        indiceCambio = 2
    elif estado != 'Sin Gestion' and estadoValidoReactData == 4:
        controlCambioPk = True
        indiceCambio = 3
    return controlCambioPk, indiceCambio

def insertarPeriodoCampanaEjecutivos(campanasEjecutivos: dict, fechaProceso):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ejecutivosExistentes = periodoCampanasReactiva(fechaProceso)
        periodoEjecutivos = convertirListaReactiva(campanasEjecutivos, ejecutivosExistentes, fechaProceso)
        if len(periodoEjecutivos) > 0:
            sql = """INSERT INTO reactiva_campanas_periodo_ejecutivos (id_ejecutivo, periodo) VALUES (?, ?);"""
            cursor.executemany(sql, periodoEjecutivos)
            db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar Periodo de Ejecutivos: %s' % (e))
    finally:
        cursor.close()
        db.close()

def insertarCampanaEjecutivos(campanasEjecutivos: dict, fechaProceso):
    try:
        db = conectorDB()
        camapanasPeriodoEjecutivos = periodoCampanasReactiva(fechaProceso)
        campanasPorPeriodo = []
        cursor = db.cursor()

        for valores in campanasEjecutivos.values():
            for polizas in valores.values():
                idEjecutivo = polizas['ID_EMPLEADO']
                if camapanasPeriodoEjecutivos.get(idEjecutivo):
                    campanasPorPeriodo += setearCampanasReactiva(polizas, camapanasPeriodoEjecutivos[idEjecutivo]['ID'])

        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INSERTAR_CAMPANAS_EJECUTIVOS': '-----------------------------------------------------' })
        campanasExistentes = CampanasPorPeriodoReactiva(fechaProceso)
        if limpiarTablaCamapanasEjecutivos(fechaProceso):
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LIMPIAR_CAMAPAÑAS_EJECUTIVOS': 'EliminarCampanaEjecutivos;Se eliminaron {0} Camapaña(s) existentes'.format(campanasExistentes)})

        if len(campanasPorPeriodo) > 0:
            sql = """INSERT INTO reactiva_campanas_ejecutivos (id_periodo_ejecutivo, numero_poliza, estado_retencion, estado_ut, in_out, certificacion, exito_repetido, estado_poliza, estado_final) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);"""
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

        sql = """DELETE FROM reactiva_campanas_ejecutivos WHERE id_periodo_ejecutivo IN (SELECT id FROM reactiva_campanas_periodo_ejecutivos WHERE periodo = ?);"""
        cursor.execute(sql, fechaProceso)
        db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al eliminar Campañas de Ejecutivos existentes: %s' % (e))
    finally:
        cursor.close()
        db.close()

def agregarCampanasPorEjecutivo(idEmpleado, pk, valoresCampanas: dict):
    
    if campanasPorEjecutivos.get(idEmpleado):
        campanasPorEjecutivos[idEmpleado].setdefault(pk, valoresCampanas)
    else:
        campanasPorEjecutivos[idEmpleado] = {pk: valoresCampanas}
    return 1

def leerArchivoReactiva(archivoEntrada, periodo, fechaInicioEntrada, fechaFinEntrada, archivoCertificacionXls, archivoComplmentoCliente):

    try:
        archivoSalida = REACTIVA_CONFIG_XLSX['SALIDA_TXT']
        encabezadoXls = REACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX']
        coordenadaEncabezado = REACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO']
        columna = REACTIVA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        complementoCliente = extraerComplementoCliente(len(LOG_PROCESO_REACTIVA), archivoComplmentoCliente)
        LOG_PROCESO_REACTIVA.update(LOG_COMPLEMENTO_CLIENTE)
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'DIVISOR_PROCESO': '-----------------------------------------------------'})

        baseCertificacion = extraerBaseCertificacion(archivoCertificacionXls)
        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'DIVISOR_PROCESO': '-----------------------------------------------------'})

        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INICIO_LECTURA_REACTIVA': 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        archivo_correcto = validarEncabezadoXlsx(hoja[coordenadaEncabezado], encabezadoXls, archivoEntrada)

        if type(archivo_correcto) is not dict:
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ENCABEZADO_REACTIVA': 'Encabezado del Archivo: %s OK' % archivoEntrada})
            campanaDescripcion = {1 : 'Inbound', 0: 'Outbound'}
            campanaIdDuplicado = dict()
            gestionReactTxt = dict()
            polizaExitoRepetido = dict()
            polizaReactTxt = dict()
            certificacionReactTxt = dict()
            dataCartolaDb = dict()

            fechaInicioPeriodo = setearFechaInput(fechaInicioEntrada)
            fechaFinPeriodo = setearFechaInput(fechaFinEntrada)
            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            ejecutivosExistentesDb = buscarRutEjecutivosDb(fechaFinMes, fechaIncioMes)
            i = 0
            cantidadCampanasValidas = 0
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INICIO_CELDAS_REACTIVA': 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Reactiva' , unit=' Fila'):

                i += 1
                if i >= 2:

                    salienteEntrada = fila[columna['LLAMADA_SALIENTE']].value
                    estado = str(fila[columna['ESTADO']].value)
                    estadoUt = fila[columna['ESTADO_ULTIMA_TAREA']].value
                    estadoRetencion = fila[columna['ESTADO_RETENCION']].value
                    numeroPoliza, numeroPolizaCertificado = formatearNumeroPoliza(fila[columna['NRO_POLIZA']].value)
                    idEmpleado = str(fila[columna['ID_EMPLEADO']].value)
                    campanaId = str(fila[columna['CAMAPAÑA_ID']].value)
                    pk = '%s_%s' % (campanaId, numeroPoliza)

                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_CREACION']])
                    fechaCierre = setearFechaCelda(fila[columna['FECHA_CIERRE']])

                    if salienteEntrada is None:
                        celdaCoordenada = setearCelda2(fila[0:columna['LLAMADA_SALIENTE']+1], len(fila[0:columna['LLAMADA_SALIENTE']])-1, i)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LLAMADA_SALIENTE': '%s;Campaña InBound/OutBound NULL;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    if fila[columna['NRO_POLIZA']].value is None:
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

                    if not ejecutivosExistentesDb.get(idEmpleado):
                        valorErroneo = fila[columna['ID_EMPLEADO']].value
                        celdaCoordenada = setearCelda2(fila[0:columna['ID_EMPLEADO']+1], len(fila[0:columna['ID_EMPLEADO']])-1, i)
                        mensaje = '%s;ID_EMPLEADO no existe en la DB;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ID_EMPLEADO': mensaje})
                        continue

                    saliente = int(salienteEntrada)
                    nombreCampana = campanaDescripcion.get(saliente)
                    exitoDuplicadoPoliza = 0

                    if saliente == inbound and fechaCreacion >= fechaIncioMes and fechaCreacion <= fechaFinMes or saliente == outbound and fechaCreacion >= fechaInicioPeriodo and fechaCreacion <= fechaFinPeriodo:

                        estadoValidoReact = validarEstadoReact(estadoRetencion, estado)
                        contactoReact = validarContactoReact(saliente, estadoRetencion, estado, estadoUt)

                        if type(contactoReact) is not int:
                            valorErroneo = str(fila[columna['ESTADO_ULTIMA_TAREA']].value)
                            celdaCoordenada = setearCelda2(fila[0:columna['ESTADO_ULTIMA_TAREA']+1], len(fila[0:columna['ESTADO_ULTIMA_TAREA']])-1, i)
                            mensaje = '%s;No existe EstadoUT;%s' % (celdaCoordenada, valorErroneo)
                            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'ERROR_ESTADO_UT': mensaje})
                            continue

                        if campanaIdDuplicado.get(campanaId):
                            pkDataGestion = campanaIdDuplicado[campanaId]['PK']
                            if contactoReact == 1 and gestionReactTxt[pkDataGestion]['CONTACTO_REACT'] == 0:
                                datosActualizados = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 0, 'ID_EMPLEADO': idEmpleado, 'ID_CAMPANA': campanaId, 'CAMPANA': nombreCampana, 'POLIZA': numeroPoliza, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                                gestionReactTxt[pkDataGestion].update(datosActualizados)
                        else:
                            campanaIdDuplicado[campanaId] = {'PK': pk}

                        if not gestionReactTxt.get(pk):
                            gestionReactTxt[pk] = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 0, 'ID_EMPLEADO': idEmpleado, 'ID_CAMPANA': campanaId, 'CAMPANA': nombreCampana, 'POLIZA': numeroPoliza, 'REPETICIONES': 1,'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                            cantidadCampanasValidas += 1
                        else:

                            controlCambioPk, indiceCambio = aprobarActualizarRegistro(estado, estadoValidoReact, contactoReact, gestionReactTxt[pk]['ESTADO_VALIDO_REACT'], gestionReactTxt[pk]['CONTACTO_REACT'])
                            if controlCambioPk:
                                datosActualizados = {'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 0, 'ID_EMPLEADO': idEmpleado, 'ID_CAMPANA': campanaId, 'CAMPANA': nombreCampana, 'POLIZA': numeroPoliza, 'REPETICIONES': gestionReactTxt[pk]['REPETICIONES'] + 1, 'FECHA_CREACION': fechaCreacion, 'FECHA_CIERRE': fechaCierre}
                                gestionReactTxt[pk].update(datosActualizados)
                                # Guardar en el log la razon del cambio
                            else:
                                gestionReactTxt[pk]['REPETICIONES'] += 1
                                continue

                        ValidacionCertificacion = 0
                        grabCertificadaReact = 0
                        if estadoValidoReact == 1:
    
                            ValidacionCertificacion = 2
                            if complementoCliente.get(int(numeroPoliza)) and complementoCliente[int(numeroPoliza)]['ESTADO_POLIZA'] == 'Vigente':
                                polizaReactTxt[numeroPoliza] = {'ESTADO_POLIZA_REACT': 1, 'NUMERO_POLIZA': numeroPoliza}

                            pkBaseCertificacion = '{0}_{1}'.format(str(numeroPoliza), idEmpleado)
                            if baseCertificacion.get(pkBaseCertificacion):

                                ejecutivoBaseCertificacion = baseCertificacion[pkBaseCertificacion]['ID_EMPLEADO']
                                if not ejecutivosExistentesDb.get(ejecutivoBaseCertificacion):
                                    valorErroneo = baseCertificacion[pkBaseCertificacion]['ID_EMPLEADO']
                                    celdaCoordenada = setearCelda2(fila[0:columna['ID_EMPLEADO']+1], len(fila[0:columna['ID_EMPLEADO']])-1, i)
                                    mensaje = '%s;EJECUTIVO_BASE_CERIFICACION no existe en la DB;%s' % (celdaCoordenada, valorErroneo)
                                    LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'EJECUTIVO_BASE_CERTIFICADO': mensaje})
                                    continue

                                campanaBaseCertificacion = campanaCanal(baseCertificacion[pkBaseCertificacion]['CANAL'])
                                fechaLlamado = baseCertificacion[pkBaseCertificacion]['FECHA_LLAMADO']


                                if saliente == 1 and campanaBaseCertificacion == 1 and idEmpleado == ejecutivoBaseCertificacion:
                                    if fechaLlamado >= fechaIncioMes and fechaLlamado <= fechaFinMes:
                                        grabCertificadaReact = 1
                                        gestionReactTxt[pk]['EXITO_REPETIDO_REACT'] = 1
                                elif saliente == 0 and campanaBaseCertificacion == 0 and idEmpleado == ejecutivoBaseCertificacion:
                                    if fechaLlamado >= fechaInicioPeriodo and fechaLlamado <= fechaFinPeriodo:
                                        grabCertificadaReact = 1
                                        gestionReactTxt[pk]['EXITO_REPETIDO_REACT'] = 1

                            certificacionReactTxt[numeroPoliza] = {'GRAB_CERTIFICADA_REACT': grabCertificadaReact, 'ID_EMPLEADO': idEmpleado, 'CANPANA': nombreCampana, 'POLIZA': numeroPoliza}

                        if grabCertificadaReact == 1:

                            ValidacionCertificacion = 1
                            if polizaExitoRepetido.get(numeroPoliza):
                                pkDataGestion = exitoRepetidoPk(numeroPoliza, polizaExitoRepetido, gestionReactTxt)

                                if pkDataGestion == 0:
                                    continue

                                exitoDuplicadoPoliza = 1
                                if saliente == 1:
                                    if fechaCreacion > gestionReactTxt[pkDataGestion]['FECHA_CREACION']:
                                        gestionReactTxt[pkDataGestion]['EXITO_REPETIDO_REACT'] = 0

                                elif saliente == 0:
                                    if type(fechaCierre) is datetime.date and type(gestionReactTxt[pkDataGestion]['FECHA_CIERRE']) is datetime.date:
                                        if fechaCierre > gestionReactTxt[pkDataGestion]['FECHA_CIERRE']:
                                            gestionReactTxt[pkDataGestion]['EXITO_REPETIDO_REACT'] = 0
                                    else:
                                        valorErroneo =  '%s-VS-%s' % (fechaCierre, gestionReactTxt[pkDataGestion]['FECHA_CIERRE'])
                                        celdaCoordenada = setearCelda2(fila[0:columna['FECHA_CIERRE']+1], len(fila[0:columna['FECHA_CIERRE']])-1, i)
                                        mensaje = '%s;FECHA_CIERRE no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FECHA_CIERRE': mensaje})
                                
                                if not polizaExitoRepetido[numeroPoliza].get(pk):
                                    polizaExitoRepetido[numeroPoliza].setdefault(pk, pk)
                            else:
                                polizaExitoRepetido[numeroPoliza] = {pk: pk}

                        estadoPoliza = complementoCliente[int(numeroPoliza)]['ESTADO_POLIZA']
                        estadoFinalDb = 0
                        if grabCertificadaReact == 1 and polizaReactTxt.get(numeroPoliza):
                            estadoFinalDb = 1

                        valoresPoliza = {'ID_EMPLEADO': idEmpleado, 'NUMERO_POLIZA': numeroPoliza, 'ESTADO_RETENCION': estadoRetencion, 'ESTAD0_UT': estadoUt, 'IN_OUT': nombreCampana, 'VALIDACION_CERTIFICACION': ValidacionCertificacion, 'EXITO_REPETIDO': exitoDuplicadoPoliza, 'ESTADO_POLIZA': estadoPoliza, 'ESTADO_FINAL': estadoFinalDb}
                        agregarCampanasPorEjecutivo(idEmpleado, pk, valoresPoliza)

            if insertarPeriodoCampanaEjecutivos(campanasPorEjecutivos, fechaIncioMes):
                if insertarCampanaEjecutivos(campanasPorEjecutivos, fechaIncioMes):
                    mensaje = 'InsertarCampanaEjecutivos;Se insertaron correctamente: %s Campaña(s)' % (cantidadCampanasValidas)
                    LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INSERTAR_CAMPANAS_EJECUTIVOS': mensaje})
                    LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'INSERTAR_CAMPANAS_EJECUTIVOS': '-----------------------------------------------------' })

            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FIN_CELDAS_REACTIVA': 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivoEntrada, len(tuple(hoja.rows)))})
            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'PROCESO_REACTIVA': 'Proceso del Archivo: %s Finalizado' % archivoEntrada})
            salidaGestionReactTxt = convertirDataReact(gestionReactTxt) 
            dataSalida = [
                {'NOMBRE_ARCHIVO': archivoSalida['GESTION']['NOMBRE_SALIDA'], 'DATA': salidaGestionReactTxt, 'ENCABEZADO': archivoSalida['GESTION']['ENCABEZADO']},
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
        return False