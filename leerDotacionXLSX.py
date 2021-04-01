from openpyxl import load_workbook
from tqdm import tqdm

from conexio_db import conectorDB
from config_xlsx import DOTACION_CONFIG_XLSX
from diccionariosDB import buscarRutEjecutivosDb
from validaciones_texto import (formatearPlataformaCRO, formatearRutGion,
                                primerDiaMes, ultimoDiaMes)

LOG_PROCESO_DOTACION = dict()

def leerArchivoDotacion(periodo):
    try:
        LOG_PROCESO_DOTACION.setdefault('INICIO_LECTURA_DOTACION', {len(LOG_PROCESO_DOTACION)+1: 'Iniciando proceso de escritura del Archivo de DOTACION'})
        encabezadoTxt = DOTACION_CONFIG_XLSX['ENCABEZADO_TXT']

        filaSalidaXls = dict()
        ejecutivosDB = buscarRutEjecutivosDb(ultimoDiaMes(periodo), primerDiaMes(periodo))
        for  idEjecutivo, valor in tqdm(iterable= ejecutivosDB.items(), total= len(ejecutivosDB), desc='Leyendo DotacionCRO' , unit=' Fila'):
            
            filaSalidaXls[idEjecutivo] = {'ID_EMPLEADO': ejecutivosDB[idEjecutivo]['ID_EMPLEADO'], 'DIRECCION': '', 'COMUNA': '', 'TELEFONO': '', 'CELULAR': '', 'FECHA_INGRESO': ejecutivosDB[idEjecutivo]['FECHA_INGRESO'], 'FECHA_NACIMIENTO': '', 'FECHA_DESVINCULACION': ejecutivosDB[idEjecutivo]['FECHA_DESVINCULACION'], 'CORREO_ELECTRONICO': '', 'RUT_JEFE': '', 'EMPRESA': 'METLIFE', 'SUCURSAL': '', 'CARGO': ejecutivosDB[idEjecutivo]['PLATAFORMA'], 'NIVEL_CARGO': '1', 'CANAL_NEGOCIO': 'MTLFCC', 'ROL_PAGO': formatearPlataformaCRO(ejecutivosDB[idEjecutivo]['PLATAFORMA'])}

        LOG_PROCESO_DOTACION.setdefault('PROCESO_DOTACION', {len(LOG_PROCESO_DOTACION)+1: 'Proceso del Archivo: DOTACION Finalizado - %s filas escritas' % len(ejecutivosDB)})
        return filaSalidaXls, encabezadoTxt
    except Exception as e:
        errorMsg = 'Error: Al escribir Archivo de DOTACION | %s' % (str(e))
        LOG_PROCESO_DOTACION.setdefault('LECTURA_ARCHIVO_DOTACION', {len(LOG_PROCESO_DOTACION)+1: errorMsg})
        return False, False
