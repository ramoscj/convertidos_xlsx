from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm
import unicodedata

from validaciones_texto import formatearRut, validarEncabezadoXlsx, primerDiaMes, ultimoDiaMes, setearCelda2, setearCelda
from config_xlsx import ASISTENCIA_CONFIG_XLSX

from diccionariosDB import buscarRutEjecutivosDb

LOG_PROCESO_ASISTENCIA = dict()

def updateEjecutivoFechaDesv(periodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ultimoDia = ultimoDiaMes(periodo)
        sql = """UPDATE ejecutivos SET fecha_desvinculacion= ? WHERE fecha_desvinculacion is NULL"""
        cursor.execute(sql, (ultimoDia,))
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al actualizar tabla de ejecutivos | %s' %e)
    finally:
        cursor.close()
        db.close()

def insertarEjecutivo(rut, nombre, nombreRrh, plataforma, periodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        primerDia = primerDiaMes(periodo)
        sql = """MERGE ejecutivos AS target
                USING (VALUES (?)) AS source (rut)
                ON (source.rut = target.rut)
                WHEN MATCHED
                THEN UPDATE
                    SET target.nombre = ?,
                        target.nombre_rrh = ?,
                        target.plataforma = ?,
                        target.fecha_desvinculacion = NULL
                WHEN NOT MATCHED
                THEN INSERT (rut, nombre, nombre_rrh, plataforma, fecha_ingreso, fecha_desvinculacion)
                    VALUES (?, ?, ?, ?, ?, NULL);"""
        valores = (rut, nombre, nombreRrh, plataforma, rut, nombre, nombreRrh, plataforma, primerDia)
        cursor.execute(sql, valores)
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al insertar ejecutivo: %s - %s' % (rut ,e))
    finally:
        cursor.close()
        db.close()

def calcularDiasHablies(columnas):
    diasDeSemana = {'LUNES': 1, 'MARTES': 2, 'MIERCOLES': 3, 'JUEVES': 4, 'VIERNES': 5}
    diasHabiles = 0
    for columna in columnas:
        for celda in columna:
            nfkd_form = unicodedata.normalize('NFKD', str(celda.value).upper())
            diaSinAcento = nfkd_form.encode('ASCII', 'ignore')
            if diasDeSemana.get(diaSinAcento.decode('utf-8')):
                diasHabiles += 1
    return diasHabiles

def leerArchivoAsistencia(archivo, periodo):
    try:
        LOG_PROCESO_ASISTENCIA.setdefault('INICIO_LECTURA_ASISTENCIA', {len(LOG_PROCESO_ASISTENCIA)+1: 'Iniciando proceso de lectura del Archivo: %s' % archivo})
        encabezadoXls = ASISTENCIA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = ASISTENCIA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = ASISTENCIA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        correlativo = 1
        archivo_correcto = validarEncabezadoXlsx(hoja['A2:C2'], encabezadoXls, archivo)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_ASISTENCIA.setdefault(len(LOG_PROCESO_ASISTENCIA)+1, {'ENCABEZADO_ASISTENCIA': 'Encabezado del Archivo: %s OK' % archivo})
            filaSalidaXls = dict()
            totalColumnas = calcularDiasHablies(hoja.iter_rows(min_row=1, min_col=5, max_row=1))
            totalFilas = len(tuple(hoja.iter_rows(min_row=3, min_col=1)))
            LOG_PROCESO_ASISTENCIA.setdefault(len(LOG_PROCESO_ASISTENCIA)+1, {'INICIO_CELDAS_ASISTENCIA': 'Iniciando lectura de Celdas del Archivo: %s' % archivo})

            updateEjecutivoFechaDesv(periodo)
            for fila in tqdm(iterable=hoja.iter_rows(min_row=3, min_col=1), total = totalFilas, desc='Leyendo AsistenciaCRO' , unit=' Fila'):
            # for fila in hoja.rows:
                diasVacaciones = 0
                if fila[columna['EJECUTIVO']].value is not None and fila[columna['RUT']].value is not None and fila[columna['PLATAFORMA']].value is not None:

                    nombreEjecutivo = str(fila[columna['EJECUTIVO']].value).lower()
                    nombreEjecutivoRrh = str(fila[columna['NOMBRE_RRH']].value)
                    rut = formatearRut(str(fila[columna['RUT']].value).upper())
                    plataforma = str(fila[columna['PLATAFORMA']].value).upper()

                    insertarEjecutivo(rut, nombreEjecutivo, nombreEjecutivoRrh, plataforma, periodo)
                    ejecutivosExistentesDb = buscarRutEjecutivosDb()

                    if not filaSalidaXls.get(rut):
                        conteoVhcAplica = 0
                        vhcAplica = 0
                        filaSalidaXls[rut] = {'CRR': correlativo}
                        for celda in range(4, totalColumnas+4):
                            if str(fila[celda].value).upper() == 'V' or str(fila[celda].value).upper() == 'VAC':
                                diasVacaciones += 1
                                conteoVhcAplica += 1
                            else:
                                conteoVhcAplica = 0
                            if conteoVhcAplica == 5:
                                vhcAplica = 1
                                conteoVhcAplica = 0
                        filaSalidaXls[rut].setdefault('VHC_MES', diasVacaciones)
                        filaSalidaXls[rut].setdefault('DIAS_HABILES_MES', totalColumnas)
                        filaSalidaXls[rut].setdefault('CARGA', vhcAplica)
                        filaSalidaXls[rut].setdefault('RUT', rut)
                        correlativo += 1
                    else:
                        errorRut = '%s - Ejecutivo duplicado: %s' % (setearCelda2(fila[columna['RUT']],0), rut)
                        LOG_PROCESO_ASISTENCIA.setdefault(len(LOG_PROCESO_ASISTENCIA)+1, {'EJECUTIVO_DUPLICADO_%s' % str(len(LOG_PROCESO_ASISTENCIA)+1): errorRut})
            LOG_PROCESO_ASISTENCIA.setdefault('FIN_CELDAS_ASISTENCIA', {len(LOG_PROCESO_ASISTENCIA)+1: 'Lectura de Celdas del Archivo: %s Finalizada - %s filas' % (archivo, len(tuple(hoja.rows)))})
            LOG_PROCESO_ASISTENCIA.setdefault('PROCESO_ASISTENCIA', {len(LOG_PROCESO_ASISTENCIA)+1: 'Proceso del Archivo: %s Finalizado' % archivo})
            return filaSalidaXls, encabezadoTxt
        else:
            LOG_PROCESO_ASISTENCIA.setdefault('ENCABEZADO_ASISTENCIA', archivo_correcto)
            raise
    except Exception as e:
        errorMsg = 'Error: %s | %s' % (archivo, e)
        LOG_PROCESO_ASISTENCIA.setdefault('LECTURA_ARCHIVO', {len(LOG_PROCESO_ASISTENCIA)+1: errorMsg})
        LOG_PROCESO_ASISTENCIA.setdefault('PROCESO_ASISTENCIA', {len(LOG_PROCESO_ASISTENCIA)+1: 'Error al procesar Archivo: %s' % archivo})
        return False, False

# leerArchivoAsistencia('INPUTS/202003_Asistencia_CRO.xlsx', '202003')
# print(LOG_PROCESO_ASISTENCIA)