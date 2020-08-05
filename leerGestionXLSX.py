from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm
import datetime
import sys;

from validaciones_texto import setearFechaInput, validarEncabezadoXlsx, setearCelda, setearFechaCelda
from diccionariosDB import buscarCamphnasDb, buscarEjecutivosDb
from config_xlsx import GESTION_CONFIG_XLSX, PATH_XLSX

from escribir_txt import salidaArchivoTxt

CAPTURADOR_ERRORES = dict()

def extraerPropietariosCro():
    # pathXlsxEntrada = 'test_xls/'
    pathXlsxEntrada = PATH_XLSX
    archivo = '%s%s.xlsx' % (pathXlsxEntrada, GESTION_CONFIG_XLSX['ENTRADA_PROPIETARIOS_XLSX'])
    try:
        encabezadoXls = GESTION_CONFIG_XLSX['ENCABEZADO_PROPIETARIOS_XLSX']
        celda = GESTION_CONFIG_XLSX['COLUMNAS_PROPIETARIOS_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        propietariosCro = dict()
        validarArchivo = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivo)
        if type(validarArchivo) is not dict:
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo PropietariosCRO' , unit=' Fila'):
                campahnaId = str(fila[celda['CAMPAÑA_ID']].value)
                if not propietariosCro.get(campahnaId):
                    nombreIBCRO = str(fila[celda['DUEÑO_NOMBRE_COMPLETO']].value).lower()
                    if fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value is None:
                        nombreNoIBCRO = str(fila[celda['CUENTA_NOMBRE_COMPLETO']].value).lower()
                    else:
                        nombreNoIBCRO = str(fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value).lower()
                    propietariosCro[campahnaId] = {'NOMBRE_IBCRO': nombreIBCRO , 'NOMBRE_NO_IBCRO': nombreNoIBCRO}
            CAPTURADOR_ERRORES.setdefault(len(CAPTURADOR_ERRORES)+1 , {'ENCABEZADO_PROPIETARIOSCRO': 'Encabezado del archivo de Prpietarios OK'})
            return propietariosCro
        else:
            CAPTURADOR_ERRORES.setdefault('ENCABEZADO_PROPIETARIOSCRO', validarArchivo)
            raise
    except Exception as e:
        errorMsg = 'Error al leer archivo: %s | %s' % (archivo, e)
        CAPTURADOR_ERRORES.setdefault(len(CAPTURADOR_ERRORES)+1 , {'LECTURA_PROPIETARIOSCRO': errorMsg})
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
        sql = "INSERT INTO codigos_cro (id, nombre) VALUES (%s, %s)"
        valores = ('NULL', nombreCampahna)
        cursor.execute(sql, valores)
        db.commit()
        buscarCampahna = buscarCamphnasDb()
        return buscarCampahna
    except Exception as e:
        raise Exception('Error al insertar campahna: %s - %s' % (nombreCampahna ,e))
    finally:
        cursor.close()
        db.close()
        
def leerArchivoGestion(archivo, periodo, fechaRangoUno, fechaRangoDos):
    try:
        # pathXlsxEntrada = PATH_XLSX
        # archivo = '%s%s' % (pathXlsxEntrada, archivo)
        encabezadoXls = GESTION_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = GESTION_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = GESTION_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        correlativo = 1

        archivo_correcto = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivo)
        if type(archivo_correcto) is not dict:
            CAPTURADOR_ERRORES.setdefault('ENCABEZADO_GESTION', {len(CAPTURADOR_ERRORES)+1: 'Encabezado del archivo de Gestion OK'})
            filaSalidaXls = dict()
            propietarioCro = extraerPropietariosCro()
            CAPTURADOR_ERRORES.setdefault('LECTURA_PROPIETARIOS', {len(CAPTURADOR_ERRORES)+1: 'Lectura del archivo de PROPIETARIOS_CRO OK'})
            campahnasExistentesDb = buscarCamphnasDb()
            ejecutivosExistentesDb = buscarEjecutivosDb()
            # rutNoExisten = []
            i = 0
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo GestiónCRO' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):
                if i >= 1:
                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_DE_CREACION']])
                    fechaUno = setearFechaInput(fechaRangoUno)
                    fechaDos = setearFechaInput(fechaRangoDos)

                    if type(fechaCreacion) is not datetime.date:
                        CAPTURADOR_ERRORES.setdefault('FECHA_CREACION', {len(CAPTURADOR_ERRORES)+1: fechaCreacion})
                        continue
                    if fechaCreacion >= fechaUno and fechaCreacion <= fechaDos:
                        campanhaId = str(fila[columna['CAMPAÑA_ID']].value)
                        nombreCampahna = str(fila[columna['NOMBRE_DE_CAMPAÑA']].value)
                        estadoUt = getEstadoUt(fila[columna['ESTADO_UT']])
                        estado = getEstado(fila[columna['ESTADO']])
                        
                        if type(estadoUt) is not int:
                            CAPTURADOR_ERRORES.setdefault('ERROR_ESTADOUT', {len(CAPTURADOR_ERRORES)+1: estadoUt})
                            continue
                        if type(estado) is not int:
                            CAPTURADOR_ERRORES.setdefault('ERROR_ESTADO', {len(CAPTURADOR_ERRORES)+1: estado})
                            continue
                        if campahnasExistentesDb.get(nombreCampahna):
                            codigoCampahnaDb = campahnasExistentesDb[nombreCampahna]['CODIGO']
                        else:
                            insertarCamphnaCro(nombreCampahna)
                            campahnasExistentesDb = buscarCamphnasDb()
                            if campahnasExistentesDb.get(nombreCampahna):
                                codigoCampahnaDb = campahnasExistentesDb[nombreCampahna]['CODIGO']

                        if nombreCampahna == 'Inbound CRO':
                            nombre_ejecutivo = propietarioCro[campanhaId]['NOMBRE_IBCRO']
                        else:
                            nombre_ejecutivo = propietarioCro[campanhaId]['NOMBRE_NO_IBCRO']
                        if ejecutivosExistentesDb.get(nombre_ejecutivo):
                            rut = ejecutivosExistentesDb[nombre_ejecutivo]['RUT']
                            filaSalidaXls[correlativo] = {'CRR': correlativo, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': rut}
                            correlativo += 1
                        else:
                            celda = str(fila[columna['CAMPAÑA_ID']])
                            errorRut = ('Celda%s - No existe Ejecutivo: %s' % (setearCelda(celda), nombre_ejecutivo))
                            CAPTURADOR_ERRORES.setdefault('EJECUTIVO_NO_EXISTE_%s' % i, {len(CAPTURADOR_ERRORES)+1: errorRut})
                    else:
                        celdaCoordenada = setearCelda(fila[columna['FECHA_DE_CREACION']])
                        errorMsg = '%s: %s no esta en el rago %s - %s' % (celdaCoordenada, fechaCreacion, fechaRangoUno, fechaRangoDos)
                        CAPTURADOR_ERRORES.setdefault('RANGO_FECHA_CREACION', {len(CAPTURADOR_ERRORES)+1: errorMsg})
                i += 1
            CAPTURADOR_ERRORES.setdefault('PROCESO_GESTION', {len(CAPTURADOR_ERRORES)+1: 'Proceso del Archivo de Gestion Finalizado'})
            return filaSalidaXls, encabezadoTxt
        else:
            CAPTURADOR_ERRORES.setdefault('ENCABEZADO_GESTION', archivo_correcto)
            CAPTURADOR_ERRORES.setdefault('PROCESO_GESTION', {len(CAPTURADOR_ERRORES)+1: 'Error al procesar Archivo de Gestion'})
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivo, e)
        CAPTURADOR_ERRORES.setdefault('LECTURA_ARCHIVO', {len(CAPTURADOR_ERRORES)+1: errorMsg})
        CAPTURADOR_ERRORES.setdefault('PROCESO_GESTION', {len(CAPTURADOR_ERRORES)+1: 'Error al procesar Archivo de Gestion'})
        return False, False

# leerArchivoGestion('Gestión CRO.xlsx', '202003', '20200101', '20200131')
# print(CAPTURADOR_ERRORES)
