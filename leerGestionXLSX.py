from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm

from validaciones_texto import setearFechaInput, validarEncabezadoXlsx
from diccionariosDB import buscarCamphnasDb, buscarEjecutivosDb

def extraerPropietariosCro():
    archivo = 'Propietarios CRO.xlsx'
    try:
        encabezadoXls = ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'FECHA DE LA ÚLTIMA MODIFICACIÓN', 'DUEÑO: NOMBRE COMPLETO', 'ASIGNADO A: NOMBRE COMPLETO', 'CUENTA: PROPIETARIO DEL CLIENTE: NOMBRE COMPLETO']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        propietariosCro = dict()
        validarArchivo = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls)
        if validarArchivo:
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo PropietariosCRO' , unit=' Fila'):
                campahnaId = str(fila[0].value)
                if not propietariosCro.get(campahnaId):
                    nombreUno = str(fila[4].value).lower()
                    if fila[5].value is None:
                        nombreDos = None
                    else:
                        nombreDos = str(fila[5].value).lower()
                    nombreTres = str(fila[6].value).lower()
                    propietariosCro[campahnaId] = {'NOMBRE_UNO': nombreUno , 'NOMBRE_DOS': nombreDos, 'NOMBRE_TRES': nombreTres}
            return propietariosCro
        else:
            raise Exception('Error el archivo de PROPIETARIOS presenta incosistencias en el encabezado')
    except Exception as e:
        raise Exception('Error al leer archivo: %s | %s' % (archivo, e))

def DefinirEstado(estadoXlsx, coordenadaFila):
    columnasError = []
    listaEstado = {'Pendiente': 1, 'Terminado con Exito': 2, 'Terminado sin Exito': 3}
    if listaEstado.get(estadoXlsx):
        return listaEstado[estadoXlsx]
    elif estadoXlsx == 'Sin Gestion':
        return 0
    else:
        resto, separador, celdaN = coordenadaFila.partition(".")
        error = 'Error estado <%s: %s' % (celdaN, estadoXlsx)
        columnasError.append(error)
        print(error)

def DefinirEstadoUt(estadoUtXlsx, coordenadaFila):
    columnasError = []
    listaEstadoUt = {'Campaña exitosa': 1, 'Teléfono ocupado': 2, 'Sin respuesta': 3, 'Campaña completada con 5 intentos': 4, 'Buzón de voz': 5, 'Llamado reprogramado': 6, 'Contacto por correo': 7, 'Teléfono apagado': 8, 'Número equivocado': 9, 'Numero invalido': 10, 'Solicita renuncia': 11, 'No quiere escuchar': 12, 'Cliente desconoce venta': 13, 'Temporalmente fuera de servicio': 14, 'Cliente vive en el extranjero': 15, 'Sin teléfono registrado': 16, 'Cliente no retenido': 17, 'No contesta': 18, 'Pendiente de envío de Póliza': 19}
    if listaEstadoUt.get(estadoUtXlsx):
        return listaEstadoUt[estadoUtXlsx]
    elif estadoUtXlsx is None:
        return 0
    else:
        resto, separador, celdaN = coordenadaFila.partition(".")
        error = 'Error estadoUT <%s: %s' % (celdaN, estadoUtXlsx)
        columnasError.append(error)
        print(error)

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
        encabezadoXls = ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO DE ÚLTIMA TAREA', 'ESTADO', 'FECHA DE LA ÚLTIMA MODIFICACIÓN', 'DUEÑO: NOMBRE COMPLETO']
        encabezadoTxt = ['CRR', 'ESTADO', 'ESTADO_UT', 'RUT', 'ID_CAMPANHA', 'CDG_CAMPANHA']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        i = 0
        j = 1
        archivo_correcto = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls)
        if archivo_correcto:
            filaSalidaXls = dict()
            propietarioCro = extraerPropietariosCro()
            campahnasExistentesDb = buscarCamphnasDb()
            ejecutivosExistentesDb = buscarEjecutivosDb()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo GestiónCRO' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):
                if str(fila[1].value).upper() != encabezadoXls[1]:
                    fechaCreacion = setearFechaInput(str(fila[1].value), str(fila[1]))
                    fechaRangoUno = setearFechaInput(str(fechaRangoUno), 'FechaUno')
                    fechaRangoDos = setearFechaInput(str(fechaRangoDos), 'FechaDos')
                    if fila[1].value is not None and fechaCreacion >= fechaRangoUno and fechaCreacion <= fechaRangoDos:
                        estadoUt = DefinirEstadoUt(fila[3].value, str(fila[3]))
                        estado = DefinirEstado(fila[4].value, str(fila[4]))
                        campanhaId = str(fila[0].value)
                        nombreCampahna = str(fila[2].value)
                        if campahnasExistentesDb.get(nombreCampahna):
                            codigoCampahnaDb = campahnasExistentesDb[nombreCampahna]['CODIGO']
                        else:
                            insertarCamphnaCro(nombreCampahna)
                            campahnasExistentesDb = buscarCamphnasDb()
                            if campahnasExistentesDb.get(nombreCampahna):
                                codigoCampahnaDb = campahnasExistentesDb[nombreCampahna]['CODIGO']
                        if propietarioCro.get(campanhaId):
                            if nombreCampahna == 'Inbound CRO':
                                if ejecutivosExistentesDb.get(propietarioCro[campanhaId]['NOMBRE_UNO']):
                                    rut = ejecutivosExistentesDb[propietarioCro[campanhaId]['NOMBRE_UNO']]['RUT']
                                    filaSalidaXls[i] = {'CRR': j, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': rut}
                                else:
                                    filaSalidaXls[i] = {'CRR': j, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': 'No existe %s Celda: - %s' % (propietarioCro[campanhaId]['NOMBRE_UNO'], str(fila[0]))}
                            else:
                                if propietarioCro[campanhaId]['NOMBRE_DOS'] is not None:
                                    if ejecutivosExistentesDb.get(propietarioCro[campanhaId]['NOMBRE_DOS']):
                                        rut = ejecutivosExistentesDb[propietarioCro[campanhaId]['NOMBRE_DOS']]['RUT']
                                        filaSalidaXls[i] = {'CRR': j, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': rut}
                                    else:
                                        filaSalidaXls[i] = {'CRR': j, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': 'No existe %s Celda: %s' % (propietarioCro[campanhaId]['NOMBRE_DOS'], str(fila[0]))}
                                else:
                                    if ejecutivosExistentesDb.get(propietarioCro[campanhaId]['NOMBRE_TRES']):
                                        rut = ejecutivosExistentesDb[propietarioCro[campanhaId]['NOMBRE_TRES']]['RUT']
                                        filaSalidaXls[i] = {'CRR': j, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': rut}
                                    else:
                                        filaSalidaXls[i] = {'CRR': j, 'ESTADO': estado, 'ESTADO_UT': estadoUt, 'ID_CAMPANHA': campanhaId, 'CDG_CAMPANHA': codigoCampahnaDb, 'RUT': 'No existe %s Celda: %s' % (propietarioCro[campanhaId]['NOMBRE_TRES'], str(fila[0]))}
                        j += 1
                        # print(j)
                    else:
                        print('Fecha: %s no esta en el rago %s - %s' % (fechaCreacion, fechaRangoUno, fechaRangoDos))
                    i += 1
            return filaSalidaXls, encabezadoTxt
        else:
            print('Error el archivo de GESTION presenta incosistencias en el encabezado')
            return False, False
    except Exception as e:
        print('Error al leer archivo: %s | %s' % (archivo, e))

# leerArchivoGestion('Gestión CRO - copia.xlsx', '202003', '20200101', '20200131')
# print(DefinirEstado('Sin Gestion', 'AA15'))
# print(buscarCamphnasDb())
# x = extraerPropietariosCro()
# if x['a5e1V000000ktfX']['NOMBRE_DOS'] is None:
#     print('estoy vacio')