from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm
import datetime
import sys;

from validaciones_texto import validarEncabezadoXlsx, setearCelda, setearFechaCelda, setearFechaInput, primerDiaMes, ultimoDiaMes, formatearFechaYM, mesSiguienteUltimoDia
from diccionariosDB import buscarCamphnasDb, buscarEjecutivosDb
from config_xlsx import GESTION_CONFIG_XLSX, PATH_XLSX

from escribir_txt import salidaArchivoTxt

LOG_PROCESO_PROACTIVA = dict()

def extraerPropietariosCro():
    # pathXlsxEntrada = 'test_xls/'
    pathXlsxEntrada = PATH_XLSX
    archivo = '%s%s.xlsx' % (pathXlsxEntrada, GESTION_CONFIG_XLSX['ENTRADA_PROPIETARIOS_XLSX'])
    LOG_PROCESO_PROACTIVA.setdefault('INICIO_LECTURA_PROPIETARIOS', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivo})
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
                if fila[celda['FECHA']].value is not None:
                    fecha = setearFechaCelda(fila[celda['FECHA']])
                else:
                    fecha = None

                if fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value is None:
                    nombreNoIBCRO = str(fila[celda['CUENTA_NOMBRE_COMPLETO']].value).lower()
                else:
                    nombreNoIBCRO = str(fila[celda['ASIGNADO_NOMBRE_COMPLETO']].value).lower()

                if not propietariosCro.get(campahnaId):
                    nombreIBCRO = str(fila[celda['DUEÑO_NOMBRE_COMPLETO']].value).lower()
                    propietariosCro[campahnaId] = {'NOMBRE_IBCRO': nombreIBCRO , 'NOMBRE_NO_IBCRO': nombreNoIBCRO, 'FECHA': fecha}
                else:
                    if fecha is not None:
                        # nombreNoIBCRO = str(fila[celda['CUENTA_NOMBRE_COMPLETO']].value).lower()
                        propietariosCro[campahnaId]['NOMBRE_NO_IBCRO'] = nombreNoIBCRO
                        propietariosCro[campahnaId]['FECHA'] = fecha
            LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1 , {'ENCABEZADO_PROPIETARIOSCRO': 'Encabezado del Archivo: %s OK' % archivo})
            LOG_PROCESO_PROACTIVA.setdefault('LECTURA_PROPIETARIOS', {len(LOG_PROCESO_PROACTIVA)+1: 'Lectura del Archivo: %s Finalizado' % archivo})
            return propietariosCro
        else:
            LOG_PROCESO_PROACTIVA.setdefault('ENCABEZADO_PROPIETARIOSCRO', validarArchivo)
            raise
    except Exception as e:
        errorMsg = 'Error al leer archivo: %s | %s' % (archivo, e)
        LOG_PROCESO_PROACTIVA.setdefault(len(LOG_PROCESO_PROACTIVA)+1 , {'LECTURA_PROPIETARIOSCRO': errorMsg})
        raise

def getEstadoContacto(celdaFila, celdaEstadoUltimaTarea):
    listaContactado = {'Terminado con exito': 1, 'Sin gestion': 0}
    listaUltimaTarea = {'Numero invalido': 0, 'Cliente retenido': 1, 'Llamado reprogramado': 1, 'Cliente no retenido': 1, 'Sin respuesta': 0, 'Buzón de voz': 0, 'Pagos al día': 0, 'Teléfono ocupado': 0, 'Teléfono apagado':	0, 'No quiere escuchar': 1, 'Contacto con el asesor':	1, 'Campaña completada con 5 intentos':	0, 'Apoyo del asesor al ejecutivo':	1, 'Número equivocado':	0, 'Pendiente respuesta cliente':	1, 'Sin gestión de cierre':	0, 'Sin teléfono registrado': 0, 'Cliente no actualizado': 0, 'Temporalmente fuera de servicio': 0, 'Carta de revocación pendiente':	1, 'Contacto por correo': 1, 'Campaña exitosa':	1, 'Solicita renuncia':	1, 'Cliente desconoce venta': 1, 'No se pudo instalar mandato':	1, 'Anulado por cambio de producto Metlife': 1, 'Cliente vive en el extranjero': 1, 'Cliente activa mandato': 1, 'Plazo previsto del producto': 0, 'Queda vigente sin pagar': 1, 'Lo está viendo con Asesor': 1}
    if listaContactado.get(celdaFila.value):
        return listaContactado[celdaFila.value]
    elif celdaFila.value == 'Pendiente' or celdaFila.value == 'Terminado sin exito':
        if listaUltimaTarea.get(celdaEstadoUltimaTarea.value):
            return listaUltimaTarea[celdaEstadoUltimaTarea.value]
        else:
            # REVISAR ESTA REGLA
            return 1
    else:
        celdaCoordenada = setearCelda(celdaFila)
        error = 'Celda%s - No existe getEstadoContacto: %s' % (celdaCoordenada, celdaFila.value)
        return error

def leerArchivoProactiva(archivoEntrada, periodo, fechaInicioEntrada, fechaFinEntrada):
    try:
        LOG_PROCESO_PROACTIVA.setdefault('INICIO_LECTURA_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivoEntrada})
        encabezadoXls = GESTION_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = GESTION_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = GESTION_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivoEntrada, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]

        archivo_correcto = validarEncabezadoXlsx(hoja['A1:G1'], encabezadoXls, archivoEntrada)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_PROACTIVA.setdefault('ENCABEZADO_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Encabezado del Archivo: %s OK' % archivoEntrada})
            filaSalidaXls = dict()

            fechaInicioPeriodo = setearFechaInput(fechaInicioEntrada)
            fechaFinPeriodo = setearFechaInput(fechaFinEntrada)
            fechaIncioMes = primerDiaMes(periodo)
            fechaFinMes = ultimoDiaMes(periodo)
            i = 0
            correlativo = 1
            LOG_PROCESO_PROACTIVA.setdefault('INICIO_CELDAS_PROACTIVA', {len(LOG_PROCESO_PROACTIVA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivoEntrada})

            filaSalidaXls = dict()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Proactiva' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):

                if i >= 1:

                    nombreCliente = str(fila[columna['NOMBRE_CLIENTE']].value)
                    fechaCreacion = setearFechaCelda(fila[columna['FECHA_CREACION']])
                    nombreCampana = str(fila[columna['NOMBRE_DE_CAMPAÑA']].value)
                    nombreEjecutivo = str(fila[columna['NOMBRE_EJECUTIVO']].value)
                    estado = str(fila[columna['ESTADO']].value)
                    fechaCierre = setearFechaCelda(fila[columna['FECHA_CIERRE']])
                    numeroPoliza = str(fila[columna['NUMERO_POLIZA']].value)
                    fechaExpiracionCoret = setearFechaCelda(fila[columna['EXPIRACION_CORET']])
                    estadoRetencion = str(fila[columna['ESTADO_RETENCION']].value)
                    estadoUltimaTarea = str(fila[columna['ESTADO_ULTIMA_TAREA']].value)

                    fechaFinMesSiguiente = mesSiguienteUltimoDia(periodo)
                    fechaIncioMes = primerDiaMes(periodo)
                    fechaFinMes = ultimoDiaMes(periodo)

                    contactado = getEstadoContacto(fila[columna['ESTADO']], fila[columna['ESTADO_ULTIMA_TAREA']])
                    retencionCobranza = False
                    retencionMandato = False
                    bonoCobranza = False
                    bonoActivacion = False

                    if type(fechaCreacion) is not datetime.date:
                        LOG_PROCESO_PROACTIVA.setdefault('FECHA_CREACION', {len(LOG_PROCESO_PROACTIVA)+1: fechaCreacion})
                        continue

                    if type(fechaCierre) is not datetime.date:
                        LOG_PROCESO_PROACTIVA.setdefault('FECHA_CIERRRE', {len(LOG_PROCESO_PROACTIVA)+1: fechaCierre})
                        continue

                    if type(fechaExpiracionCoret) is not datetime.date:
                        LOG_PROCESO_PROACTIVA.setdefault('FECHA_CIERRRE', {len(LOG_PROCESO_PROACTIVA)+1: fechaExpiracionCoret})
                        continue

                    if type(contactado) is not int:
                        LOG_PROCESO_PROACTIVA.setdefault('ERROR_ESTADO', {len(LOG_PROCESO_PROACTIVA)+1: contactado})
                        continue

                    if type(fechaFinMesSiguiente) is not datetime.date:
                        LOG_PROCESO_PROACTIVA.setdefault('MES_SIGUIENTE', {len(LOG_PROCESO_PROACTIVA)+1: fechaFinMesSiguiente})
                        continue

                    if estado != 'Sin Gestion' and fechaCierre >= fechaIncioMes and fechaCierre <= fechaFinMes or estado == 'Sin Gestion' and fechaExpiracionCoret >= fechaIncioMes or fechaExpiracionCoret <= fechaFinMesSiguiente:

                        if nombreCampana == 'CO RET - Cobranza' and estado == 'Terminado con Exito':
                            retencionCobranza = True
                            if estadoRetencion == 'Mantiene su producto' or estadoRetencion == 'Realiza pago en línea':
                                bonoCobranza = True
                            elif estadoRetencion == 'Realiza Activación PAC/PAT':
                                bonoCobranza = True
                                bonoActivacion = True
                        elif nombreCampana == 'CO RET - Fallo en Instalacion de Mandato' and estado == 'Terminado con Exito':
                            retencionMandato = True
                            if estadoRetencion == 'Realiza pago en línea':
                                bonoCobranza = True
                            elif estadoRetencion == 'Realiza Activación PAC/PAT':
                                bonoActivacion = True
                            elif estadoRetencion == 'Mantiene su producto':
                                bonoCobranza = True
                                bonoActivacion = True
                        else:
                            continue

                        
                    else:
                        continue


                i += 1
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

# leerArchivoGestion('test_xls/Gestión CRO - copia.xlsx', '202003')
# print(LOG_PROCESO_GESTION)