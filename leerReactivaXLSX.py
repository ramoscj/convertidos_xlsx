from openpyxl import load_workbook
from tqdm import tqdm
import datetime
from conexio_db import conectorDB

from validaciones_texto import validarEncabezadoXlsx, setearCelda, setearFechaCelda, primerDiaMes, ultimoDiaMes, setearFechaInput, formatearFechaMesAnterior, setearCelda2, formatearIdCliente
from diccionariosDB import buscarEjecutivosDb, buscarPolizasReliquidar, buscarPolizasReliquidarAll
from config_xlsx import REACTIVA_CONFIG_XLSX, PATH_XLSX, listaEstadoContactado

from escribir_txt import salidaArchivoTxt, salidaLogTxt

from complementoCliente import extraerComplementoCliente, LOG_COMPLEMENTO_CLIENTE

LOG_PROCESO_REACTIVA = dict()

def extraerBaseCertificacion(complementoCliente: dict):
    pathXlsxEntrada = 'test_xls/REACTIVA/'
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
                        valorErroneo = str(fila[celda['FECHA_LLAMADO']].value)
                        celdaCoordenada = setearCelda2(fila[0:celda['FECHA_LLAMADO']+1], len(fila[0:celda['FECHA_LLAMADO']])-1, i)
                        mensaje = '%s;FECHA_LLAMADO no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FECHA_LLAMADO': mensaje})
                        continue

                    if not complementoCliente.get(numeroPoliza) or complementoCliente[numeroPoliza]['ESTADO_POLIZA'] != 'Vigente':
                        celdaCoordenada = setearCelda2(fila[celda['NRO_POLIZA']],0)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'POLIZA_ESTADO_NOK': '%s;Numero de poliza no cumple ESTADO en complementoCliente;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    baseCertificado[numeroPoliza] = {'NRO_POLIZA': str(numeroPoliza), 'FECHA_LLAMADO': fechaLlamado, 'EJECUTIVO': ejecutivo, 'CANAL': canal, 'TIPO_CERTIFICACION': tipoCertificacion}

            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'LECTURA_BASE_CERTIFICACION': 'Lectura del Archivo: %s Finalizado - %s Filas' % (archivo, len(tuple(hoja.rows)))})
            return baseCertificado
        else:
            LOG_PROCESO_REACTIVA.setdefault('ENCABEZADO_BASE_CERTIFICACION', validarArchivo)
    except Exception as e:
        errorMsg = 'Error al leer archivo;%s | %s' % (archivo, e)
        LOG_PROCESO_REACTIVA.setdefault('LECTURA_BASE_CERTIFICACION' , errorMsg)
        raise

def leerArchivoReactiva(archivoEntrada, periodo, fechaInicioEntrada, fechaFinEntrada):
    listaEstado = {'No gestionable': 1, 'Pendiente': 2, 'Terminado con exito': 3, 'Terminado sin exito': 4}

    try:
        LOG_PROCESO_REACTIVA.setdefault('INICIO_LECTURA_PROACTIVA', {len(LOG_PROCESO_REACTIVA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        encabezadoXls = REACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = REACTIVA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = REACTIVA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        archivo_correcto = validarEncabezadoXlsx(hoja['A1:N1'], encabezadoXls, archivoEntrada)
        complementoCliente = extraerComplementoCliente(len(LOG_PROCESO_REACTIVA))
        baseCertificacion = extraerBaseCertificacion(complementoCliente)
        print(baseCertificacion)
        LOG_PROCESO_REACTIVA.update(LOG_COMPLEMENTO_CLIENTE)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_REACTIVA.setdefault('ENCABEZADO_REACTIVA', {len(LOG_PROCESO_REACTIVA)+1: 'Encabezado del Archivo: %s OK' % archivoEntrada})
            filaSalidaXls = dict()

            fechaInicioPeriodo = setearFechaInput(fechaInicioEntrada)
            fechaFinPeriodo = setearFechaInput(fechaFinEntrada)
            periodoMesAnterior = formatearFechaMesAnterior(periodo)
            fechaIncioMes = primerDiaMes(periodoMesAnterior.strftime("%Y%m"))
            fechaFinMes = ultimoDiaMes(periodoMesAnterior.strftime("%Y%m"))
            i = 0
            correlativo = 1
            LOG_PROCESO_REACTIVA.setdefault('INICIO_CELDAS_REACTIVA', {len(LOG_PROCESO_REACTIVA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Reactiva' , unit=' Fila'):

                i += 1
                if i >= 2:

                    saliente = str(fila[columna['LLAMADA_SALIENTE']].value)
                    estado = str(fila[columna['ESTADO']].value)
                    estadoUt = fila[columna['ESTADO_ULTIMA_TAREA']].value
                    estadoRetencion = fila[columna['ESTADO_RETENCION']].value
                    numeroPoliza = str(fila[columna['NRO_POLIZA']].value)
                    nombreCliente = str(fila[columna['NOMBRE_CLIENTE']].value)
                    nombreEjecutivo = str(fila[columna['NOMBRE_EJECUTIVO']].value)
                    campanaId = str(fila[columna['CAMAPAÑA_ID']].value)
                    idCliente = formatearIdCliente(nombreCliente)
                    pk = '%s_%s' % (campanaId, numeroPoliza)

                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_CREACION']])

                    if numeroPoliza is None:
                        celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'POLIZA_NULO': '%s;Numero de poliza NULL;%s' % (celdaCoordenada, numeroPoliza)})
                        continue

                    if type(fechaCreacion) is not datetime.date:
                        valorErroneo = str(fila[columna['FECHA_CREACION']].value)
                        celdaCoordenada = setearCelda2(fila[0:columna['FECHA_CREACION']+1], len(fila[0:columna['FECHA_CREACION']])-1, i)
                        mensaje = '%s;FECHA_CREACION no es una fecha valida;%s' % (celdaCoordenada, valorErroneo)
                        LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'FECHA_CREACION': mensaje})
                        continue

                    if saliente == 0 and fechaCreacion >= fechaIncioMes and fechaCreacion <= fechaFinMes or saliente == 1 and fechaCreacion >= fechaInicioPeriodo and fechaCreacion <= fechaFinPeriodo:

                        if estadoRetencion == 'Mantiene su producto' or estadoRetencion is None and estado == 'Terminado con Exito':
                            estadoValidoReact = 'Terminado con Exito'
                        elif estadoRetencion is None:
                            estadoValidoReact = estado
                        elif estadoRetencion != 'Mantiene su producto':
                            estadoValidoReact = 'No gestionable'

                        saliente = 0
                        if saliente == 0:
                            contactoReact = 1
                        elif saliente == 1:
                            if estadoRetencion == 'Mantiene su producto':
                                contactoReact = 1
                            elif estadoRetencion != 'Mantiene su producto':
                                if estadoUt is None:
                                    if estadoRetencion == 'Mantiene su producto':
                                        contactoReact = 1
                                    elif estadoRetencion is None:
                                        if estado == 'Terminado con Exito':
                                            contactoReact = 1
                                elif listaEstadoContactado.get(estadoUt):
                                    contactoReact = 1

                        if filaSalidaXls.get(pk):
                            celdaCoordenada = setearCelda2(fila[0:columna['NRO_POLIZA']+1], len(fila[0:columna['NRO_POLIZA']])-1, i)
                            mensaje = '%s;Poliza duplicada;%s' % (celdaCoordenada, pk)
                            LOG_PROCESO_REACTIVA.setdefault(len(LOG_PROCESO_REACTIVA)+1, {'POLIZA_DUPLICADA': mensaje})
                            continue

                        # Certificar retencion
                        if baseCertificacion.get(numeroPoliza):
                            if baseCertificacion[numeroPoliza]['EJECUTIVO']:
                                fechaLlamado = baseCertificacion[numeroPoliza]['FECHA_LLAMADO']
                                canal = baseCertificacion[numeroPoliza]['CANAL']
                                if saliente == 0 and canal.upper() == 'INBOUND CORET' and fechaLlamado >= fechaIncioMes and fechaLlamado <= fechaFinMes:
                                    pass
                                elif saliente == 1 and canal.upper() != 'INBOUND CORET' and fechaLlamado >= fechaInicioPeriodo and fechaLlamado <= fechaFinPeriodo:
                                    pass

                        filaSalidaXls[pk] = {'CRR': correlativo, 'ESTADO_VALIDO_REACT': estadoValidoReact, 'CONTACTO_REACT': contactoReact, 'EXITO_REPETIDO_REACT': 0, 'GRAB_CERTIFICADA_REACT': 0, 'RUT': 0, 'ID_CAMPANA': campanaId, 'CAMPANA': 0, 'POLIZA': numeroPoliza, 'ID_CLIENTE': idCliente}
                    correlativo += 1

            LOG_PROCESO_REACTIVA.setdefault('FIN_CELDAS_REACTIVA', {len(LOG_PROCESO_REACTIVA)+1: 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivoEntrada, len(tuple(hoja.rows)))})
            LOG_PROCESO_REACTIVA.setdefault('PROCESO_REACTIVA', {len(LOG_PROCESO_REACTIVA)+1: 'Proceso del Archivo: %s Finalizado' % archivoEntrada})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_REACTIVA.setdefault('ENCABEZADO_REACTIVA', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivoEntrada, e)
        LOG_PROCESO_REACTIVA.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_REACTIVA)+1: errorMsg})
        LOG_PROCESO_REACTIVA.setdefault('PROCESO_REACTIVA', {len(LOG_PROCESO_REACTIVA)+1: 'Error al procesar Archivo: %s' % archivoEntrada})
        return False, False
        # raise

x, y = leerArchivoReactiva('test_xls/REACTIVA/Gestion Reactiva.xlsx', '202009', '20201001', '202010031')
salidaLogTxt('test_xls/REACTIVA/reactiva.log', LOG_PROCESO_REACTIVA)
print(salidaArchivoTxt('test_xls/REACTIVA/%s%s.txt' % (REACTIVA_CONFIG_XLSX['SALIDA_TXT'],'202010'), x, y))
