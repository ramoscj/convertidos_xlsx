import unicodedata

from openpyxl import load_workbook
from tqdm import tqdm

from conexio_db import conectorDB
from config_xlsx import ASISTENCIA_CONFIG_XLSX
from validaciones_texto import (formatearRut, primerDiaMes, setearCelda,
                                setearCelda2, ultimoDiaMes,
                                validarEncabezadoXlsx)

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

def insertarEjecutivo(idEmpleado, plataforma, periodo):
    try:
        db = conectorDB()
        cursor = db.cursor()
        primerDia = primerDiaMes(periodo)
        sql = """MERGE ejecutivos AS target
                USING (VALUES (?)) AS source (id_empleado)
                ON (source.id_empleado = target.id_empleado)
                WHEN MATCHED
                THEN UPDATE
                    SET target.plataforma = ?,
                        target.fecha_desvinculacion = NULL
                WHEN NOT MATCHED
                THEN INSERT (id_empleado, plataforma, fecha_ingreso, fecha_desvinculacion)
                    VALUES (?, ?, ?, NULL);"""
        valores = (idEmpleado, plataforma, idEmpleado, plataforma, primerDia)
        cursor.execute(sql, valores)
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al insertar ejecutivo: %s - %s' % (idEmpleado ,e))
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
        coordenadaEcabezado = ASISTENCIA_CONFIG_XLSX['COORDENADA_ENCABEZADO']
        columnasAdicionales = ASISTENCIA_CONFIG_XLSX['COLUMNAS_ADICIONALES']
        xls = load_workbook(archivo, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        correlativo = 1
        archivo_correcto = validarEncabezadoXlsx(hoja[coordenadaEcabezado], encabezadoXls, archivo)
        if type(archivo_correcto) is not dict:
            LOG_PROCESO_ASISTENCIA.setdefault(len(LOG_PROCESO_ASISTENCIA)+1, {'ENCABEZADO_ASISTENCIA': 'Encabezado del Archivo: %s OK' % archivo})
            filaSalidaXls = dict()
            totalColumnas = calcularDiasHablies(hoja.iter_rows(min_row=1, min_col=2, max_row=1))
            totalFilas = len(tuple(hoja.iter_rows(min_row=3, min_col=1)))
            LOG_PROCESO_ASISTENCIA.setdefault(len(LOG_PROCESO_ASISTENCIA)+1, {'INICIO_CELDAS_ASISTENCIA': 'Iniciando lectura de Celdas del Archivo: %s' % archivo})

            updateEjecutivoFechaDesv(periodo)
            for fila in tqdm(iterable=hoja.iter_rows(min_row=3, min_col=1), total = totalFilas, desc='Leyendo AsistenciaCRO' , unit=' Fila'):
            # for fila in hoja.rows:
                diasVacaciones = 0
                ausentismoMes = 1
                if fila[columna['ID_EMPLEADO']].value is not None and fila[columna['PLATAFORMA']].value is not None:

                    idEmpleado = fila[columna['ID_EMPLEADO']].value
                    plataforma = str(fila[columna['PLATAFORMA']].value).upper()

                    insertarEjecutivo(idEmpleado, plataforma, periodo)

                    if not filaSalidaXls.get(idEmpleado):
                        conteoVhcAplica = 0
                        vhcAplica = 0
                        filaSalidaXls[idEmpleado] = {'CRR': correlativo}
                        for celda in range(2, totalColumnas + columnasAdicionales):
                            if str(fila[celda].value).upper() == 'V' or str(fila[celda].value).upper() == 'VAC':
                                diasVacaciones += 1
                                conteoVhcAplica += 1
                            else:
                                conteoVhcAplica = 0

                            if conteoVhcAplica == 5:
                                vhcAplica = 1
                                conteoVhcAplica = 0

                            if type(fila[celda].value) is int and fila[celda].value == 1:
                                ausentismoMes = 0
                        filaSalidaXls[idEmpleado].setdefault('VHC_MES', diasVacaciones)
                        filaSalidaXls[idEmpleado].setdefault('DIAS_HABILES_MES', totalColumnas)
                        filaSalidaXls[idEmpleado].setdefault('CARGA', vhcAplica)
                        filaSalidaXls[idEmpleado].setdefault('AUSENTISMO_MES', ausentismoMes)
                        filaSalidaXls[idEmpleado].setdefault('ID_EMPLEADO', idEmpleado)
                        correlativo += 1
                    else:
                        errorRut = '%s - Ejecutivo duplicado: %s' % (setearCelda2(fila[columna['ID_EMPLEADO']],0), idEmpleado)
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

# leerArchivoAsistencia('PROACTIVA/INPUTS/202012_Asistencia Plataforma.xlsx', '202012')
# print(LOG_PROCESO_ASISTENCIA)
