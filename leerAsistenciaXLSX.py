from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx, primerDiaMes, ultimoDiaMes, setearCelda
from config_xlsx import ASISTENCIA_CONFIG_XLSX

LOG_PROCESO_ASISTENCIA = dict()

def updateEjecutivoFechaDesv(periodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        ultimoDia = ultimoDiaMes(periodo)
        sql = """UPDATE ejecutivos SET fecha_desvinculacion=%s WHERE fecha_desvinculacion is NULL"""
        cursor.execute(sql, (ultimoDia,))
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al actualizar tabla de ejecutivos | %s' %e)
    finally:
        cursor.close()
        db.close()

def insertarEjecutivo(rut, nombre, plataforma, periodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        primerDia = primerDiaMes(periodo)
        sql = "INSERT INTO ejecutivos (id, rut, nombre, plataforma, fecha_ingreso, fecha_desvinculacion) VALUES (NULL, %s, %s, %s, %s, NULL) ON DUPLICATE KEY UPDATE nombre=%s, plataforma=%s, fecha_desvinculacion=NULL"
        valores = (rut, nombre, plataforma, primerDia, nombre, plataforma)
        cursor.execute(sql, valores)
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al insertar ejecutivo: %s - %s' % (rut ,e))
    finally:
        cursor.close()
        db.close()

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
            LOG_PROCESO_ASISTENCIA.setdefault('ENCABEZADO_ASISTENCIA', {len(LOG_PROCESO_ASISTENCIA)+1: 'Encabezado del Archivo: %s OK' % archivo})
            filaSalidaXls = dict()
            totalColumnas = len(tuple(hoja.iter_cols(min_row=3, min_col=1)))
            totalFilas = len(tuple(hoja.iter_rows(min_row=3, min_col=1)))
            LOG_PROCESO_ASISTENCIA.setdefault('INICIO_CELDAS_GESTION', {len(LOG_PROCESO_ASISTENCIA)+1: 'Iniciando lectura de Celdas del Archivo: %s' % archivo})

            updateEjecutivoFechaDesv(periodo)
            for fila in tqdm(iterable=hoja.iter_rows(min_row=3, min_col=1), total = totalFilas, desc='Leyendo AsistenciaCRO' , unit=' Fila'):
            # for fila in hoja.rows:
                diasVacaciones = 0
                if fila[columna['EJECUTIVA']].value is not None and fila[columna['RUT']].value is not None and fila[columna['PLATAFORMA']].value is not None:

                    nombreEjecutivo = str(fila[columna['EJECUTIVA']].value).lower()
                    rut = formatearRut(str(fila[columna['RUT']].value))
                    plataforma = str(fila[columna['PLATAFORMA']].value).upper()

                    insertarEjecutivo(rut, nombreEjecutivo, plataforma, periodo)
                    if not filaSalidaXls.get(rut):
                        conteoVhcAplica = 0
                        vhcAplica = 0
                        filaSalidaXls[rut] = {'CRR': correlativo}
                        for celda in range(3, totalColumnas):
                            if str(fila[celda].value).upper() == 'V' or str(fila[celda].value).upper() == 'VAC':
                                diasVacaciones += 1
                                conteoVhcAplica += 1
                            else:
                                conteoVhcAplica = 0
                            if conteoVhcAplica == 5:
                                vhcAplica = 1
                                conteoVhcAplica = 0
                        filaSalidaXls[rut].setdefault('VHC_MES', diasVacaciones)
                        filaSalidaXls[rut].setdefault('DIAS_HABILES_MES', totalColumnas - 3)
                        filaSalidaXls[rut].setdefault('CARGA', vhcAplica)
                        filaSalidaXls[rut].setdefault('RUT', rut)
                        correlativo += 1
                    else:
                        errorRut = 'Celda%s - Ejecutivo duplicado: %s' % (setearCelda(fila[columna['RUT']]), rut)
                        LOG_PROCESO_ASISTENCIA.setdefault('EJECUTIVO_DUPLICADO_%s' % correlativo, {len(LOG_PROCESO_ASISTENCIA)+1: errorRut})
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