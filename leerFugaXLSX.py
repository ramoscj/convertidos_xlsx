from openpyxl import load_workbook
from tqdm import tqdm
import datetime

from validaciones_texto import formatearRut, validarEncabezadoXlsx, setearCelda, validaFechaCelda
from config_xlsx import FUGA_CONFIG_XLSX, PATH_XLSX
from diccionariosDB import buscarRutEjecutivosDb

LOG_PROCESO_FUGA = dict()

def validarFugaStock(correlativo, tipo, lpattrCodStat, considerarFuga, rut, unidad):
    rutNuevoFuga = dict()
    rutNuevoFuga["CRR"] = correlativo
    if tipo == 'FUGA' and considerarFuga != 'NO':
        rutNuevoFuga["FUGA"] = 1
    else:
        rutNuevoFuga["FUGA"] = 0
    if  lpattrCodStat != 'NVIG' and considerarFuga != 'NO':
        rutNuevoFuga["STOCK"] = 1
    else:
        rutNuevoFuga["STOCK"] = 0
    rutNuevoFuga["RUT"] = rut
    rutNuevoFuga["UNIDAD"] = unidad
    return rutNuevoFuga

def existeRut(tipo, lpattrCodStat, considerarFuga, rutExistenteFuga):
    if tipo == 'FUGA' and  considerarFuga != 'NO':
        rutExistenteFuga["FUGA"] += 1
    elif  lpattrCodStat != 'NVIG' and considerarFuga != 'NO':
        rutExistenteFuga["STOCK"] += 1
    return rutExistenteFuga

def leerArchivoFuga(archivo, periodo):
    try:
        LOG_PROCESO_FUGA.setdefault('INICIO_LECTURA_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivo})
        encabezadoXls = FUGA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoFugaTxt = FUGA_CONFIG_XLSX['ENCABEZADO_FUGA_TXT']
        # encabezadoStockTxt = FUGA_CONFIG_XLSX['ENCABEZADO_STOCK_TXT']
        columna = FUGA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        archivo_correcto = validarEncabezadoXlsx(hoja['A1:M1'], encabezadoXls, archivo)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_FUGA.setdefault('ENCABEZADO_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Encabezado del Archivo: %s OK' % archivo})
            i = 0
            correlativo = 1
            filaSalidaFugaXls = dict()
            filaSalidaStockXls = dict()
            ejecutivosExistentesDb = buscarRutEjecutivosDb()
            LOG_PROCESO_FUGA.setdefault('INICIO_CELDAS_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivo})
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo FugaCRO' , unit=' Fila'):

                if i >= 1:
                    fechaLpattrs = validaFechaCelda(fila[columna['LPATTR_PER_RES']])
                    if type(fechaLpattrs) is str:
                            LOG_PROCESO_FUGA.setdefault('FECHA_CREACION', {len(LOG_PROCESO_FUGA)+1: fechaLpattrs})
                            continue
                    if periodo == str(fechaLpattrs.value) and fila[columna['RUT_CRO']].value is not None:

                        rut = formatearRut(str(fila[columna['RUT_CRO']].value).upper())
                        tipo = str(fila[columna['TIPO']].value).upper()
                        considerarFuga = str(fila[columna['CONSIDERAR_FUGA']].value).upper()
                        lpattrCodStat = str(fila[columna['LPATTR_COD_STAT']].value).upper()

                        if ejecutivosExistentesDb.get(rut):
                            unidad = ejecutivosExistentesDb[rut]['PLATAFORMA']
                            if filaSalidaFugaXls.get(rut):
                                filaSalidaFugaXls[rut] = existeRut(tipo, lpattrCodStat, considerarFuga, filaSalidaFugaXls[rut])
                            else:
                                filaSalidaFugaXls[rut] = validarFugaStock(correlativo, tipo, lpattrCodStat, considerarFuga, rut, unidad)
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
            # return filaSalidaFugaXls, filaSalidaStockXls, encabezadoFugaTxt, encabezadoStockTxt
            return filaSalidaFugaXls, encabezadoFugaTxt
        else:
            LOG_PROCESO_FUGA.setdefault('ENCABEZADO_FUGA', archivo_correcto)
            raise Exception('Error en enbezado de archivo: %s' % archivo)
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivo, e)
        LOG_PROCESO_FUGA.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_FUGA)+1: errorMsg})
        LOG_PROCESO_FUGA.setdefault('PROCESO_FUGA', {len(LOG_PROCESO_FUGA)+1: 'Error al procesar Archivo: %s' % archivo})
        # print(e)
        return False, False

# x,y = leerArchivoFuga('INPUTS/202007_Fuga_Agencia.xlsx', '202007')
# print(LOG_PROCESO_FUGA)
# print(y)