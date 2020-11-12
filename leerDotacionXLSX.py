from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm

from validaciones_texto import primerDiaMes, ultimoDiaMes, formatearRutGion, formatearPlataformaCRO
from config_xlsx import DOTACION_CONFIG_XLSX

from diccionariosDB import buscarEjecutivosAllDb

LOG_PROCESO_DOTACION = dict()

def leerArchivoDotacion(periodo):
    try:
        LOG_PROCESO_DOTACION.setdefault('INICIO_DOTACION', {len(LOG_PROCESO_DOTACION)+1: '-----------------------------------------------------'})
        LOG_PROCESO_DOTACION.setdefault('INICIO_LECTURA_DOTACION', {len(LOG_PROCESO_DOTACION)+1: 'Iniciando proceso de escritura del Archivo de GESTION'})
        encabezadoTxt = DOTACION_CONFIG_XLSX['ENCABEZADO_TXT']

        filaSalidaXls = dict()
        ejecutivosDB = buscarEjecutivosAllDb(ultimoDiaMes(periodo), primerDiaMes(periodo))
        # print(ejecutivosDB)
        for  rut, valor in tqdm(iterable= ejecutivosDB.items(), total= len(ejecutivosDB), desc='Leyendo AsistenciaCRO' , unit=' Fila'):
        # # for fila in hoja.rows:

            filaSalidaXls[rut] = {'RUT': formatearRutGion(ejecutivosDB[rut]['RUT'])}
            filaSalidaXls[rut].setdefault('NOMBRES', str(ejecutivosDB[rut]['NOMBRES']))
            filaSalidaXls[rut].setdefault('APELLIDO_PATERNO', str(ejecutivosDB[rut]['APELLIDO_PATERNO']))
            filaSalidaXls[rut].setdefault('APELLIDO_MATERNO', str(ejecutivosDB[rut]['APELLIDO_MATERNO']))
            filaSalidaXls[rut].setdefault('DIRECCION', '')
            filaSalidaXls[rut].setdefault('COMUNA', '')
            filaSalidaXls[rut].setdefault('TELEFONO', '')
            filaSalidaXls[rut].setdefault('CELULAR', '')
            filaSalidaXls[rut].setdefault('FECHA_INGRESO', ejecutivosDB[rut]['FECHA_INGRESO'])
            filaSalidaXls[rut].setdefault('FECHA_NACIMIENTO', '')
            filaSalidaXls[rut].setdefault('FECHA_DESVINCULACION', ejecutivosDB[rut]['FECHA_DESVINCULACION'])
            filaSalidaXls[rut].setdefault('CORREO_ELECTRONICO', '')
            filaSalidaXls[rut].setdefault('RUT_JEFE', '')
            filaSalidaXls[rut].setdefault('EMPRESA', 'METLIFE')
            filaSalidaXls[rut].setdefault('SUCURSAL', '')
            filaSalidaXls[rut].setdefault('CARGO', ejecutivosDB[rut]['PLATAFORMA'])
            filaSalidaXls[rut].setdefault('NIVEL_CARGO', '1')
            filaSalidaXls[rut].setdefault('CANAL_NEGOCIO', 'MTLFCC')
            filaSalidaXls[rut].setdefault('ROL_PAGO', formatearPlataformaCRO(ejecutivosDB[rut]['PLATAFORMA']))

        LOG_PROCESO_DOTACION.setdefault('PROCESO_DOTACION', {len(LOG_PROCESO_DOTACION)+1: 'Proceso del Archivo: GESTION Finalizado'})
        return filaSalidaXls, encabezadoTxt
    except Exception as e:
        errorMsg = 'Error: Archivo de GESTION | %s' % (e)
        LOG_PROCESO_DOTACION.setdefault('LECTURA_ARCHIVO_DOTACION', {len(LOG_PROCESO_DOTACION)+1: errorMsg})
        LOG_PROCESO_DOTACION.setdefault('PROCESO_DOTACION', {len(LOG_PROCESO_DOTACION)+1: 'Error al procesar Archivo: GESTION'})
        return False, False
        # raise e