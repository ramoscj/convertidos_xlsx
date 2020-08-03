from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm

from validaciones_texto import formatearRut

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
        encabezadoXls = []
        encabezadoTxt = ['CRR', 'VHC_MES', 'DIAS_HABILES_MES', 'VHC_APLICA', 'RUT']
        xls = load_workbook(archivo, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        j = 1
        filaSalidaXls = dict()
        totalColumnas = len(tuple(hoja.iter_cols(min_row=1, min_col=1)))
        totalFilas = len(tuple(hoja.iter_rows(min_row=3, min_col=1)))
        for fila in tqdm(iterable=hoja.iter_rows(min_row=3, min_col=1), total = totalFilas, desc='Leyendo AsistenciaCRO' , unit=' Fila'):
        # for fila in hoja.iter_rows(min_row=3, min_col=1):
            diasVacaciones = 0
            if fila[0].value is not None and fila[1].value is not None and fila[2].value is not None:

                rut = formatearRut(str(fila[1].value))
                nombreEjecutivo = str(fila[0].value).lower()
                plataforma = str(fila[2].value).upper()

                insertarEjecutivo(rut, nombreEjecutivo, plataforma)

                if not filaSalidaXls.get(rut):
                    conteoCarga = 0
                    cargaTxt = 0
                    filaSalidaXls[rut] = {'CRR': j}
                    for columna in range(3, totalColumnas):
                        if str(fila[columna].value).upper() == 'V' or str(fila[columna].value).upper() == 'VAC':
                            diasVacaciones += 1
                            conteoCarga += 1
                        else:
                            conteoCarga = 0
                        if conteoCarga == 5:
                            cargaTxt = 1
                            conteoCarga = 0
                    filaSalidaXls[rut].setdefault('VHC_MES', diasVacaciones)
                    filaSalidaXls[rut].setdefault('DIAS_HABILES_MES', totalColumnas - 3)
                    filaSalidaXls[rut].setdefault('CARGA', cargaTxt)
                    filaSalidaXls[rut].setdefault('RUT', rut)
                    j += 1
                else:
                    raise Exception('El RUT: %s esta duplicado en la DATA' % (fila[1].value))
        return filaSalidaXls, encabezadoTxt
    except Exception as e:
        raise Exception('Error al leer archivo: %s | %s' % (archivo, e))