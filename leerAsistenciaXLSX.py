from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx
from config_xlsx import ASISTENCIA_CONFIG_XLSX

def insertarEjecutivo(rut, nombre, plataforma):
    try:
        db = conectorDB()
        cursor = db.cursor()
        sql = "INSERT INTO ejecutivos (id, rut, nombre, plataforma) VALUES (NULL,%s, %s, %s) ON DUPLICATE KEY UPDATE nombre=%s, plataforma=%s"
        valores = (rut, nombre, plataforma, nombre, plataforma)
        cursor.execute(sql, valores)
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al insertar ejecutivo: %s - %s' % (rut ,e))

def leerArchivoAsistencia(archivo, periodo):
    try:
        encabezadoXls = ASISTENCIA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = ASISTENCIA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = ASISTENCIA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        j = 1
        archivo_correcto = validarEncabezadoXlsx(hoja['A2:C2'], encabezadoXls)
        if archivo_correcto:
            filaSalidaXls = dict()
            totalColumnas = len(tuple(hoja.iter_cols(min_row=1, min_col=1)))
            totalFilas = len(tuple(hoja.iter_rows(min_row=3, min_col=1)))
            for fila in tqdm(iterable=hoja.iter_rows(min_row=3, min_col=1), total = totalFilas, desc='Leyendo AsistenciaCRO' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):
                diasVacaciones = 0
                if fila[columna['EJECUTIVA']].value is not None and fila[columna['RUT']].value is not None and fila[columna['PLATAFORMA']].value is not None:

                    nombreEjecutivo = str(fila[columna['EJECUTIVA']].value).lower()
                    rut = formatearRut(str(fila[columna['RUT']].value))
                    plataforma = str(fila[columna['PLATAFORMA']].value).upper()

                    insertarEjecutivo(rut, nombreEjecutivo, plataforma)
                    if not filaSalidaXls.get(rut):
                        conteoVhcAplica = 0
                        vhcAplica = 0
                        filaSalidaXls[rut] = {'CRR': j}
                        for columna in range(3, totalColumnas):
                            if str(fila[columna].value).upper() == 'V' or str(fila[columna].value).upper() == 'VAC':
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
                        j += 1
                    else:
                        raise Exception('El RUT: %s esta duplicado en la DATA' % (fila[columna['RUT']].value))
            return filaSalidaXls, encabezadoTxt
        else:
            print('Error el archivo de ASISTENCIA presenta incosistencias en el encabezado')
    except Exception as e:
        raise Exception('Error al leer archivo: %s | %s' % (archivo, e))