from openpyxl import load_workbook
from tqdm import tqdm

from config_xlsx import CALIDAD_CONFIG_XLSX
from diccionariosDB import buscarRutEjecutivosDb
from validaciones_texto import (formatearRut, setearCelda2,
                                validarEncabezadoXlsx)

LOG_PROCESO_CALIDAD = dict()

def leerArchivoCalidad(archivo, periodo):
    try:
        LOG_PROCESO_CALIDAD.setdefault('INICIO_LECTURA', {len(LOG_PROCESO_CALIDAD)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivo})
        encabezadoXls = CALIDAD_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = CALIDAD_CONFIG_XLSX['ENCABEZADO_TXT']
        celda = CALIDAD_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        archivo_correcto = validarEncabezadoXlsx(hoja['A2:D2'], encabezadoXls, archivo)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_CALIDAD.setdefault(len(LOG_PROCESO_CALIDAD)+1, {'ENCABEZADO_ARCHIVO': 'Encabezado del Archivo: %s OK' % archivo})
            i = 0
            correlativo = 1
            filaSalidaXls = dict()
            ejecutivosExistentesDb = buscarRutEjecutivosDb()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo Calidad' , unit=' Fila'):
            # for fila in hoja.rows:
                if i >= 2 and fila[celda['RUT']].value is not None:
                    rut = formatearRut(str(fila[celda['RUT']].value).upper())
                    if ejecutivosExistentesDb.get(rut):
                        calidad = int(float(fila[celda['CALIDAD']].value)*100)
                        filaSalidaXls[rut] = {'CRR': correlativo, 'CALIDAD': calidad, 'RUT': rut}
                        correlativo += 1
                    else:
                        errorRut = '%s;No existe Ejecutivo;%s' % (setearCelda2(fila[celda['RUT']], 0), rut)
                        LOG_PROCESO_CALIDAD.setdefault(len(LOG_PROCESO_CALIDAD)+1, {'EJECUTIVO_NO_EXISTE_%s' % i: errorRut})
                i += 1
            LOG_PROCESO_CALIDAD.setdefault(len(LOG_PROCESO_CALIDAD)+1, {'FIN_LECTURA_ARCHIVO': 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivo, len(tuple(hoja.rows)))})
            LOG_PROCESO_CALIDAD.setdefault(len(LOG_PROCESO_CALIDAD)+1, {'FIN_PROCESO': 'Proceso del Archivo: %s Finalizado' % archivo})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_CALIDAD.setdefault('ENCABEZADO_ARCHIVO', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivo, e)
        LOG_PROCESO_CALIDAD.setdefault(len(LOG_PROCESO_CALIDAD)+1, {'LECTURA_ARCHIVO_ERROR': errorMsg})
        LOG_PROCESO_CALIDAD.setdefault(len(LOG_PROCESO_CALIDAD)+1, {'PROCESO_ARCHIVO': 'Error al procesar Archivo: %s' % archivo})
        return False, False

# leerArchivoCalidad('INPUTS/202005_Calidad_CRO.xlsx', '202005')
# print(LOG_PROCESO_CALIDAD)
