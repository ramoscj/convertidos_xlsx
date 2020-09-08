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
    # pathXlsxEntrada = 'test_xls/'
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
        if type(validarArchivo) is not dict:
            for fila in tqdm(iterable=hoja.rows, total= len(tuple(hoja.rows)), desc= 'Leyendo PropietariosCRO' , unit=' Fila'):
                campahnaId = str(fila[celda['CAMPAÑA_ID']].value)
                if not propietariosCro.get(campahnaId):
                    nombreIBCRO = str(fila[celda['DUEÑO_NOMBRE_COMPLETO']].value).lower()
                    if fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value is None:
                        nombreNoIBCRO = str(fila[celda['CUENTA_NOMBRE_COMPLETO']].value).lower()
                    else:
                        nombreNoIBCRO = str(fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value).lower()
                    propietariosCro[campahnaId] = {'NOMBRE_IBCRO': nombreIBCRO , 'NOMBRE_NO_IBCRO': nombreNoIBCRO}
            LOG_PROCESO_GESTION.setdefault(len(LOG_PROCESO_GESTION)+1 , {'ENCABEZADO_PROPIETARIOSCRO': 'Encabezado del Archivo: %s OK' % archivo})
            LOG_PROCESO_GESTION.setdefault('LECTURA_PROPIETARIOS', {len(LOG_PROCESO_GESTION)+1: 'Lectura del Archivo: %s Finalizado' % archivo})
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
        error = 'Celda%s - No existe estado: %s' % (celdaCoordenada, celdaFila.value)
        return error

def getEstadoUt(celdaFila):
    columnasError = []
    listaEstadoUt = {'Campaña exitosa': 1, 'Teléfono ocupado': 2, 'Sin respuesta': 3, 'Campaña completada con 5 intentos': 4, 'Buzón de voz': 5, 'Llamado reprogramado': 6, 'Contacto por correo': 7, 'Teléfono apagado': 8, 'Número equivocado': 9, 'Numero invalido': 10, 'Solicita renuncia': 11, 'No quiere escuchar': 12, 'Cliente desconoce venta': 13, 'Temporalmente fuera de servicio': 14, 'Cliente vive en el extranjero': 15, 'Sin teléfono registrado': 16, 'Cliente no retenido': 17, 'No contesta': 18, 'Pendiente de envío de Póliza': 19}
    if listaEstadoUt.get(celdaFila.value):
        return listaEstadoUt[celdaFila.value]
    elif celdaFila.value is None:
        return 0
    else:
        celdaCoordenada = setearCelda(celdaFila)
        error = 'Celda%s - No existe estadoUt: %s' % (celdaCoordenada, celdaFila.value)
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
        LOG_PROCESO_GESTION.setdefault('INICIO_LECTURA_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        encabezadoXls = GESTION_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = GESTION_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = GESTION_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        archivo_correcto = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivoEntrada)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_GESTION.setdefault('ENCABEZADO_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Encabezado del Archivo: %s OK' % archivoEntrada})
            filaSalidaXls = dict()
            propietarioCro = extraerPropietariosCro()
            campahnasExistentesDb = buscarCamphnasDb()
            ejecutivosExistentesDb = buscarEjecutivosDb()

            fechaInicioPeriodo = setearFechaInput(fechaInicioEntrada)
            fechaFinPeriodo = setearFechaInput(fechaFinEntrada)
            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            i = 0
            correlativo = 1
            LOG_PROCESO_GESTION.setdefault('INICIO_CELDAS_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo GestionCRO' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):

                if i >= 1:
                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_DE_CREACION']])
                    fechaUltimaModificacion = setearFechaCelda(fila[columna['FECHA_ULTIMA_MODF']])

                    estado = getEstado(fila[columna['ESTADO']])
                    nombreCampahna = str(fila[columna['NOMBRE_DE_CAMPAÑA']].value)
                    campanhaId = str(fila[columna['CAMPAÑA_ID']].value)
                    estadoUt = getEstadoUt(fila[columna['ESTADO_UT']])

                    # test
                    nombreCompletoGestion = str(fila[columna['NOMBRE_COMPLETO']].value)

                    if type(fechaUltimaModificacion) is not datetime.date:
                        LOG_PROCESO_GESTION.setdefault('FECHA_ULTIMA_MODF', {len(LOG_PROCESO_GESTION)+1: fechaUltimaModificacion})
                        continue

                    if type(fechaCreacion) is not datetime.date:
                        LOG_PROCESO_GESTION.setdefault('FECHA_CREACION', {len(LOG_PROCESO_GESTION)+1: fechaCreacion})
                        continue

                    if type(estadoUt) is not int:
                        LOG_PROCESO_GESTION.setdefault('ERROR_ESTADOUT', {len(LOG_PROCESO_GESTION)+1: estadoUt})
                        continue
                    if type(estado) is not int:
                        LOG_PROCESO_GESTION.setdefault('ERROR_ESTADO', {len(LOG_PROCESO_GESTION)+1: estado})
                        continue

                    if nombreCampahna == 'Inbound CRO':
                        if fechaUltimaModificacion >= fechaIncioMes and fechaUltimaModificacion <= fechaFinMes:
                            if estado == 0:
                                continue
                            else:
                                nombre_ejecutivo = propietarioCro[campanhaId]['NOMBRE_IBCRO']
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
                        filaSalidaXls[correlativo] = {'CRR': correlativo, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CAMPANA': nombreCampahna[0:30], 'RUT': rut, 'NOMBRE_GESTION': nombreCompletoGestion, 'NOMBRE_PROPIETARIO': nombre_ejecutivo}
                        correlativo += 1
                    else:
                        # errorRut = 'Celda%s - No existe Ejecutivo: %s' % (setearCelda(fila[columna['CAMPAÑA_ID']]), nombre_ejecutivo)
                        # LOG_PROCESO_GESTION.setdefault('EJECUTIVO_NO_EXISTE_%s' % i, {len(LOG_PROCESO_GESTION)+1: errorRut})
                        rut = 'SIN RUT'
                        filaSalidaXls[correlativo] = {'CRR': correlativo, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CAMPANA': nombreCampahna[0:30], 'RUT': rut, 'NOMBRE_GESTION': nombreCompletoGestion, 'NOMBRE_PROPIETARIO': nombre_ejecutivo}
                        correlativo += 1
                    # else:
                    #     celdaCoordenada = setearCelda(fila[columna['FECHA_DE_CREACION']])
                    #     errorMsg = '%s: %s no esta en el rago %s - %s' % (celdaCoordenada, fechaCreacion, fechaRangoUno, fechaRangoDos)
                    #     LOG_PROCESO_GESTION.setdefault('RANGO_FECHA_CREACION', {len(LOG_PROCESO_GESTION)+1: errorMsg})
                i += 1
            LOG_PROCESO_GESTION.setdefault('FIN_CELDAS_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivoEntrada, len(tuple(hoja.rows)))})
            LOG_PROCESO_GESTION.setdefault('PROCESO_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Proceso del Archivo: %s Finalizado' % archivoEntrada})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_GESTION.setdefault('ENCABEZADO_GESTION', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivoEntrada, e)
        LOG_PROCESO_GESTION.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_GESTION)+1: errorMsg})
        LOG_PROCESO_GESTION.setdefault('PROCESO_GESTION', {len(LOG_PROCESO_GESTION)+1: 'Error al procesar Archivo: %s' % archivoEntrada})
        return False, False

# leerArchivoGestion('test_xls/Gestión CRO - copia.xlsx', '202003')
# print(LOG_PROCESO_GESTION)