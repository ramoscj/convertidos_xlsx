from openpyxl import load_workbook
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx, setearCelda
from diccionariosDB import buscarRutEjecutivosDb
from config_xlsx import CALIDAD_CONFIG_XLSX

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
            LOG_PROCESO_CALIDAD.setdefault('ENCABEZADO_ARCHIVO', {len(LOG_PROCESO_CALIDAD)+1: 'Encabezado del Archivo: %s OK' % archivo})
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
                        errorRut = 'Celda%s - No existe Ejecutivo: %s' % (setearCelda(fila[celda['RUT']]), rut)
                        LOG_PROCESO_CALIDAD.setdefault('EJECUTIVO_NO_EXISTE_%s' % i, {len(LOG_PROCESO_CALIDAD)+1: errorRut})
                i += 1
            LOG_PROCESO_CALIDAD.setdefault('FIN_LECTURA_ARCHIVO', {len(LOG_PROCESO_CALIDAD)+1: 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivo, len(tuple(hoja.rows)))})
            LOG_PROCESO_CALIDAD.setdefault('FIN_PROCESO', {len(LOG_PROCESO_CALIDAD)+1: 'Proceso del Archivo: %s Finalizado' % archivo})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_CALIDAD.setdefault('ENCABEZADO_ARCHIVO', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivo, e)
        LOG_PROCESO_CALIDAD.setdefault('LECTURA_ARCHIVO_ERROR', {len(LOG_PROCESO_CALIDAD)+1: errorMsg})
        LOG_PROCESO_CALIDAD.setdefault('PROCESO_ARCHIVO', {len(LOG_PROCESO_CALIDAD)+1: 'Error al procesar Archivo: %s' % archivo})
        return False, False

# leerArchivoCalidad('INPUTS/202005_Calidad_CRO.xlsx', '202005')
# print(LOG_PROCESO_CALIDAD)