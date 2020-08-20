from openpyxl import load_workbook
from tqdm import tqdm
import datetime

from validaciones_texto import formatearRut, validarEncabezadoXlsx, setearCelda, validaFechaCelda
from config_xlsx import FUGA_CONFIG_XLSX, PATH_XLSX
from diccionariosDB import buscarRutEjecutivosDb

LOG_PROCESO_FUGA = dict()

def validarFugaStock(correlativo, tipo, lpattrCodStat, considerarFuga, rut):
    rutNuevo = dict()
    rutNuevo["CRR"] = correlativo
    if tipo == 'FUGA' and considerarFuga != 'NO':
        rutNuevo["FUGA"] = 1
    else:
        rutNuevo["FUGA"] = 0
    if  lpattrCodStat != 'NVIG' or considerarFuga == 'NO':
        rutNuevo["STOCK"] = 1
    else:
        rutNuevo["STOCK"] = 0
    rutNuevo["RUT"] = rut
    return rutNuevo

def existeRut(tipo, lpattrCodStat, considerarFuga, rutExistente):
    if tipo == 'FUGA' and  considerarFuga != 'NO':
        rutExistente["FUGA"] += 1
    elif  lpattrCodStat != 'NVIG' or considerarFuga == 'NO':
        rutExistente["STOCK"] += 1
    return rutExistente

def leerArchivoFuga(archivo, periodo):
    try:
        LOG_PROCESO_FUGA.setdefault('INICIO_LECTURA_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivo})
        encabezadoXls = FUGA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = FUGA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = FUGA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        archivo_correcto = validarEncabezadoXlsx(hoja['A1:M1'], encabezadoXls, archivo)
        LOG_PROCESO_FUGA.setdefault('ENCABEZADO_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Encabezado del Archivo: %s OK' % archivo})
        if type(archivo_correcto) is not dict:
            i = 0
            correlativo = 1
            filaSalidaXls = dict()
            x = []
            ejecutivosExistentesDb = buscarRutEjecutivosDb()
            LOG_PROCESO_FUGA.setdefault('INICIO_CELDAS_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivo})
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo FugaCRO' , unit=' Fila'):

                if i >= 1:
                    fechaLpattrs = validaFechaCelda(fila[columna['LPATTR_PER_RES']])
                    if type(fechaLpattrs) is str:
                            LOG_PROCESO_FUGA.setdefault('FECHA_CREACION', {len(LOG_PROCESO_FUGA)+1: fechaLpattrs})
                            continue
                    if periodo == str(fechaLpattrs.value) and fila[columna['RUT_CRO']].value is not None:
                        
                        rut = formatearRut(str(fila[columna['RUT_CRO']].value))
                        tipo = str(fila[columna['TIPO']].value).upper()
                        considerarFuga = str(fila[columna['CONSIDERAR_FUGA']].value).upper()
                        lpattrCodStat = str(fila[columna['LPATTR_COD_STAT']].value).upper()

                        if ejecutivosExistentesDb.get(rut):

                            if filaSalidaXls.get(rut):
                                filaSalidaXls[rut] = existeRut(tipo, lpattrCodStat, considerarFuga, filaSalidaXls[rut])
                            else:
                                filaSalidaXls[rut] = validarFugaStock(correlativo, tipo, lpattrCodStat, considerarFuga, rut)
                                correlativo += 1
                        else:
                            errorRut = 'Celda%s - No existe Ejecutivo: %s' % (setearCelda(fila[columna['RUT_CRO']]), rut)
                            LOG_PROCESO_FUGA.setdefault('EJECUTIVO_NO_EXISTE_%s' % i, {len(LOG_PROCESO_FUGA)+1: errorRut})
                    # else:
                    #     errorfecha = 'Celda%s - Error en fecha: %s' % (setearCelda(fechaLpattrs), str(fechaLpattrs.value))
                    #     LOG_PROCESO_FUGA.setdefault('ERROR_FECHA_%s' % i, {len(LOG_PROCESO_FUGA)+1: errorfecha})
                i += 1

            LOG_PROCESO_FUGA.setdefault('FIN_CELDAS_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivo, len(tuple(hoja.rows)))})
            LOG_PROCESO_FUGA.setdefault('PROCESO_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Proceso del Archivo: %s Finalizado' % archivo})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_FUGA.setdefault('ENCABEZADO_FUGA', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivo, e)
        LOG_PROCESO_FUGA.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_FUGA)+1: errorMsg})
        LOG_PROCESO_FUGA.setdefault('PROCESO_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Error al procesar Archivo: %s' % archivo})
        return False, False

# print(leerArchivoFuga('INPUTS/201911_Fuga_Agencia - copia.xlsx', '201911'))