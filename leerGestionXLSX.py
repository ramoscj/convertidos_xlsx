from openpyxl import load_workbook
from conexio_db import conectorDB
from validaciones_texto import setearFechaInput
from tqdm import tqdm

def leerArchivoAsistencia(archivo, periodo, fechaRangoUno, fechaRangoDos):
    try:
        encabezadoXls = ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO DE ÚLTIMA TAREA', 'ESTADO', 'FECHA DE LA ÚLTIMA MODIFICACIÓN', 'DUEÑO: NOMBRE COMPLETO']
        encabezadoTxt = ['CRR', 'ESTADO', 'ESTADO_UT', 'RUT', 'ID_CAMPANHA', 'CDG_CAMPANHA']
        listaEstado = {'Sin Gestion': 0, 'Pendiente': 1, 'Terminado con Exito': 2, 'Terminado sin Exito': 3}
        listaEstadoUt = {'Campaña exitosa': 1, 'Teléfono ocupado': 2, 'Sin respuesta': 3, 'Campaña completada con 5 intentos': 4, 'Buzón de voz': 5, 'Llamado reprogramado': 6, 'Contacto por correo': 7, 'Teléfono apagado': 8, 'Número equivocado': 9, 'Numero invalido': 10, 'Solicita renuncia': 11, 'No quiere escuchar': 12, 'Cliente desconoce venta': 13, 'Temporalmente fuera de servicio': 14, 'Cliente vive en el extranjero': 15, 'Sin teléfono registrado': 16, 'Cliente no retenido': 17, 'No contesta': 18, 'Pendiente de envío de Póliza': 19}
        xls = load_workbook(archivo, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        j = 1
        archivo_correcto = False
        for fila in hoja['A1:G1']:
            for celda in fila:
                if str(celda.value).upper() == encabezadoXls[i]:
                    archivo_correcto = True
                else:
                    print('Error columna: %s' % i)
                    archivo_correcto = False
                i += 1
        if archivo_correcto:
            filaSalidaXls = dict()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo DATA' , unit=' Fila'):
            # for fila in hoja.iter_rows(min_row=3, min_col=1):
                fechaCreacion = setearFechaInput(str(fila[1].value))
                if fila[1].value is not None and fechaCreacion >= fechaRangoUno and fechaCreacion <= fechaRangoDos:
                    estado = str(fila[4].value)
                    estadoUt = str(fila[3].value)
                    if listaEstado.get(estado):
                        if listaEstadoUt.get(estadoUt):
                            nombreEjecutivo = str(fila[0].value).lower()
                            plataforma = str(fila[2].value).upper()
                        elif str(fila[3].value) is None:
                            pass
                        else:
                            print('El ESTADO_UT: %s no existe' % estadoUt)
                    else:
                        print('El ESTADO: %s no existe' % estado)

                    # insertarEjecutivo(rut, nombreEjecutivo, plataforma)
            return filaSalidaXls, encabezadoTxt
        else:
            print('Error el archivo de GESTION presenta incosistencias en el encabezado')
            return False, False
    except Exception as e:
        print('Error al leer archivo: %s | %s' % (archivo, e))

