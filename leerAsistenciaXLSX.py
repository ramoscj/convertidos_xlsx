from openpyxl import load_workbook
from conexio_db import conectorDB
from tqdm import tqdm

def insertarEjecutivo(rut, nombre, plataforma):
    try:
        db = conectorDB()
        cursor = db.cursor()
        sql = "INSERT INTO ejecutivos (rut, nombre, plataforma) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE nombre=%s, plataforma=%s"
        valores = (rut, nombre, plataforma, nombre, plataforma)
        cursor.execute(sql, valores)
        db.commit()
        return True
    except Exception as e:
        raise Exception('Error al insertar ejecutivo: %s - %s' % (rut ,e))

def leerArchivoAsistencia(archivo, periodo):
    try:
        encabezadoXls = []
        encabezadoTxt = ['CRR', 'VHC_MES', 'DIAS_HABILES_MES', 'RUT']
        xls = load_workbook(archivo, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        j = 1
        filaSalidaXls = dict()
        insertDB = 0
        totalColumnas = len(tuple(hoja.iter_cols(min_row=1, min_col=1)))
        totalFilas = len(tuple(hoja.iter_rows(min_row=3, min_col=1)))
        for fila in tqdm(iterable=hoja.iter_rows(min_row=3, min_col=1), total = totalFilas, desc='Leyendo DATA' , unit=' Fila'):
        # for fila in hoja.iter_rows(min_row=3, min_col=1):
            diasVacaciones = 0
            if fila[0].value is not None and fila[1].value is not None and fila[2].value is not None:

                rutMantisa, separador, dv = str(fila[1].value).partition("-")
                rut = '%s%s' % (rutMantisa, dv)
                nombreEjecutivo = str(fila[0].value).lower()
                plataforma = str(fila[2].value).upper()

                insertarEjecutivo(rut, nombreEjecutivo, plataforma)

                if not filaSalidaXls.get(rut):
                    filaSalidaXls[rut] = {'CRR': j}
                    for columna in range(3, totalColumnas):
                        if str(fila[columna].value).upper() == 'V' or str(fila[columna].value).upper() == 'VAC':
                            diasVacaciones += 1
                    filaSalidaXls[rut].setdefault('VACACIONES', diasVacaciones)
                    filaSalidaXls[rut].setdefault('DIAS_HABILES', totalColumnas - 3)
                    filaSalidaXls[rut].setdefault('RUT', rut)
                    j += 1
                else:
                    raise Exception('Error el RUT: %s esta duplicado en la DATA' % (fila[1].value))
        return filaSalidaXls, encabezadoTxt
    except Exception as e:
        print('Error al leer archivo: %s | %s' % (archivo, e))