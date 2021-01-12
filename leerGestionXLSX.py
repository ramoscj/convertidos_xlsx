from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm
import datetime
import sys;

from validaciones_texto import validarEncabezadoXlsx, setearCelda, setearFechaCelda, setearFechaInput, primerDiaMes, ultimoDiaMes
from diccionariosDB import buscarCamphnasDb, buscarEjecutivosDb
from config_xlsx import GESTION_CONFIG_XLSX, PATH_XLSX

from escribir_txt import salidaArchivoTxt

LOG_PROCESO_GESTION = dict()

def extraerPropietariosCro():
    # pathXlsxEntrada = 'test_xls/TEST_GESTION/'
    pathXlsxEntrada = PATH_XLSX
    archivo = '%s%s.xlsx' % (pathXlsxEntrada, GESTION_CONFIG_XLSX['ENTRADA_PROPIETARIOS_XLSX'])
    LOG_PROCESO_GESTION.setdefault('INICIO_LECTURA_PROPIETARIOS', {len(LOG_PROCESO_GESTION)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivo})
    try:
        encabezadoXls = GESTION_CONFIG_XLSX['ENCABEZADO_PROPIETARIOS_XLSX']
        celda = GESTION_CONFIG_XLSX['COLUMNAS_PROPIETARIOS_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        propietariosCro = dict()
        validarArchivo = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivo)
        ejecutivosExistentesDb = buscarEjecutivosDb()
        if type(validarArchivo) is not dict:
            for fila in tqdm(iterable=hoja.rows, total= len(tuple(hoja.rows)), desc= 'Leyendo PropietariosCRO' , unit=' Fila'):

                campahnaId = str(fila[celda['CAMPAÑA_ID']].value)
                fecha = None
                if fila[celda['FECHA']].value is not None:
                    fecha = setearFechaCelda(fila[celda['FECHA']])

                nombreNoIBCRO = str(fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value).lower()
                if not ejecutivosExistentesDb.get(nombreNoIBCRO):
                    nombreNoIBCRO = str(fila[celda['CUENTA_NOMBRE_COMPLETO']].value).lower()

                if not propietariosCro.get(campahnaId):
                    propietariosCro[campahnaId] = {'NOMBRE_NO_IBCRO': nombreNoIBCRO, 'FECHA': fecha}
                else:
                    if fecha is not None:
                        if propietariosCro[campahnaId]['FECHA'] is None:
                            propietariosCro[campahnaId]['NOMBRE_NO_IBCRO'] = nombreNoIBCRO
                            propietariosCro[campahnaId]['FECHA'] = fecha
                        else:
                            if fecha > propietariosCro[campahnaId]['FECHA']:
                                propietariosCro[campahnaId]['NOMBRE_NO_IBCRO'] = nombreNoIBCRO
                                propietariosCro[campahnaId]['FECHA'] = fecha

            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1 , {'ENCABEZADO_PROPIETARIOSCRO': 'Encabezado del Archivo: %s OK' % archivo})
            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'LECTURA_PROPIETARIOS': 'Lectura del Archivo: %s Finalizado' % archivo})
            return propietariosCro
        else:
            LOG_PROCESO_GESTION.setdefault('ENCABEZADO_PROPIETARIOSCRO', validarArchivo)
            raise
    except Exception as e:
        errorMsg = 'Error al leer archivo: %s | %s' % (archivo, e)
        LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1 , {'LECTURA_PROPIETARIOSCRO': errorMsg})
        raise

def getEstado(celdaFila):
    listaEstado = {'Pendiente': 1, 'Terminado con Exito': 2, 'Terminado sin Exito': 3}
    if listaEstado.get(celdaFila.value):
        return listaEstado[celdaFila.value]
    elif celdaFila.value == 'Sin Gestion':
        return 0
    else:
        celdaCoordenada = setearCelda(celdaFila)
        error = 'Celda%s;No existe estado;%s' % (celdaCoordenada, celdaFila.value)
        return error

def getEstadoUt(celdaFila):
    listaEstadoUt = {'Campaña exitosa': 1, 'Teléfono ocupado': 2, 'Sin respuesta': 3, 'Campaña completada con 5 intentos': 4, 'Buzón de voz': 5, 'Llamado reprogramado': 6, 'Contacto por correo': 7, 'Teléfono apagado': 8, 'Número equivocado': 9, 'Numero invalido': 10, 'Solicita renuncia': 11, 'No quiere escuchar': 12, 'Cliente desconoce venta': 13, 'Temporalmente fuera de servicio': 14, 'Cliente vive en el extranjero': 15, 'Sin teléfono registrado': 16, 'Cliente no retenido': 17, 'No contesta': 18, 'Pendiente de envío de Póliza': 19}
    if listaEstadoUt.get(celdaFila.value):
        return listaEstadoUt[celdaFila.value]
    elif celdaFila.value is None:
        return 0
    else:
        celdaCoordenada = setearCelda(celdaFila)
        error = 'Celda%s;No existe estadoUt;%s' % (celdaCoordenada, celdaFila.value)
        return error

def insertarCamphnaCro(nombreCampahna):
    try:
        db = conectorDB()
        cursor = db.cursor()
        sql = "INSERT INTO codigos_cro (nombre) VALUES (?)"
        cursor.execute(sql, (nombreCampahna))
        db.commit()
        buscarCampahna = buscarCamphnasDb()
        return buscarCampahna
    except Exception as e:
        raise Exception('Error al insertar campahna: %s - %s' % (nombreCampahna ,e))
    finally:
        cursor.close()
        db.close()

def leerArchivoGestion(archivoEntrada, periodo, fechaInicioEntrada, fechaFinEntrada):
    try:
        LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'INICIO_LECTURA_GESTION': 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        encabezadoXls = GESTION_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = GESTION_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = GESTION_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        archivo_correcto = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivoEntrada)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'ENCABEZADO_GESTION': 'Encabezado del Archivo: %s OK' % archivoEntrada})
            filaSalidaXls = dict()
            ejecutivosNoExisten = dict()
            propietarioCro = extraerPropietariosCro()
            print(propietarioCro)
            campahnasExistentesDb = buscarCamphnasDb()
            ejecutivosExistentesDb = buscarEjecutivosDb()

            fechaInicioPeriodo = setearFechaInput(fechaInicioEntrada)
            fechaFinPeriodo = setearFechaInput(fechaFinEntrada)
            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            i = 0
            correlativo = 1
            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'INICIO_CELDAS_GESTION': 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo GestionCRO' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):

                if i >= 1:
                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_DE_CREACION']])

                    estado = getEstado(fila[columna['ESTADO']])
                    nombreCampahna = str(fila[columna['NOMBRE_DE_CAMPAÑA']].value)
                    campanhaId = str(fila[columna['CAMPAÑA_ID']].value)
                    estadoUt = getEstadoUt(fila[columna['ESTADO_UT']])

                    if type(fechaCreacion) is not datetime.date:
                        LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'FECHA_CREACION': fechaCreacion})
                        continue

                    if type(estadoUt) is not int:
                        LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'ERROR_ESTADOUT': estadoUt})
                        continue
                    if type(estado) is not int:
                        LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'ERROR_ESTADO': estado})
                        continue

                    if nombreCampahna == 'Inbound CRO':
                        if estado != 0:
                            if propietarioCro.get(campanhaId):
                                fechaUltimaModificacion = propietarioCro[campanhaId]['FECHA']
                                if fechaUltimaModificacion is None:
                                    errorCampana = 'Celda%s;FECHA NULL Archivo Propietario;%s' % (setearCelda(fila[columna['CAMPAÑA_ID']]), campanhaId)
                                    LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'FECHA_PROPIETARIO_NONE': errorCampana})
                                    continue
                                if fechaUltimaModificacion >= fechaIncioMes and fechaUltimaModificacion <= fechaFinMes:
                                    nombre_ejecutivo = propietarioCro[campanhaId]['NOMBRE_NO_IBCRO']
                                else:
                                    continue
                            else:
                                errorCampana = 'Celda%s;No existe Campaña en PROPIETARIOS_CRO;%s' % (setearCelda(fila[columna['CAMPAÑA_ID']]), campanhaId)
                                LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'CAMPANA_NO_EXISTE': errorCampana})
                                continue
                        else:
                            continue
                    else:
                        if fechaCreacion < fechaInicioPeriodo or fechaCreacion > fechaFinPeriodo:
                            continue
                        nombre_ejecutivo = propietarioCro[campanhaId]['NOMBRE_NO_IBCRO']

                    if not campahnasExistentesDb.get(nombreCampahna):
                        insertarCamphnaCro(nombreCampahna)
                        campahnasExistentesDb = buscarCamphnasDb()

                    if ejecutivosExistentesDb.get(nombre_ejecutivo):
                        rut = ejecutivosExistentesDb[nombre_ejecutivo]['RUT']
                        filaSalidaXls[correlativo] = {'CRR': correlativo, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CAMPANA': nombreCampahna[0:30], 'RUT': rut}
                        correlativo += 1
                    else:
                        errorRut = 'Celda%s;No existe Ejecutivo;%s' % (setearCelda(fila[columna['CAMPAÑA_ID']]), nombre_ejecutivo)
                        LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'EJECUTIVO_NO_EXISTE': errorRut})
                i += 1
            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'FIN_CELDAS_GESTION': 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivoEntrada, len(tuple(hoja.rows)))})
            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1, {'PROCESO_GESTION': 'Proceso del Archivo: %s Finalizado' % archivoEntrada})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_GESTION.setdefault('ENCABEZADO_GESTION', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivoEntrada, e)
        LOG_PROCESO_GESTION.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_GESTION)+1: errorMsg})
        LOG_PROCESO_GESTION.setdefault('PROCESO_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Error al procesar Archivo: %s' % archivoEntrada})
        return False, False

# x,y = leerArchivoGestion('test_xls/TEST_GESTION/Gestión CRO.xlsx', '202011', '20201023', '20201125')
# print(salidaArchivoTxt('test_xls/TEST_GESTION/test.txt', x, y))
# print(LOG_PROCESO_GESTION)

# nombre = 'catalina berrios'
# print(nombre.count('berrios'))