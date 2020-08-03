from openpyxl import load_workbook
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx
from diccionariosDB import buscarRutEjecutivosDb

def leerArchivoCampanhasEsp(archivo, periodo):
    try:
        encabezadoXls = ['EJECUTIVA', 'RUT', 'PLATAFORMA', 'CANTIDAD GESTIONES CAMPAÑAS ESPECIALES']
        encabezadoTxt = ['CRR', 'NUMERO_GESTIONES', 'RUT']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        archivo_correcto = validarEncabezadoXlsx(hoja['A2:D2'], encabezadoXls)
        if archivo_correcto:
            i = 0
            j = 1
            filaSalidaXls = dict()
            ejecutivosExistentesDb = buscarRutEjecutivosDb()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo CamapañasEspeciales' , unit=' Fila'):
            # for fila in hoja.rows:
                if i >= 2 and fila[1].value is not None:
                    rut = formatearRut(str(fila[1].value))
                    if ejecutivosExistentesDb.get(rut):
                        numeroGestiones = fila[3].value
                        filaSalidaXls[rut] = {'CRR': j, 'NUMERO_GESTIONES': numeroGestiones, 'RUT': rut}
                        j += 1
                    else:
                        filaSalidaXls[rut] = {'CRR': j, 'NUMERO_GESTIONES': numeroGestiones, 'RUT': 'No existe %s' % rut}
                i += 1
            return filaSalidaXls, encabezadoTxt
        else:
            raise Exception('Incosistencias en el encabezado')
    except Exception as e:
        raise Exception('Error al leer archivo: %s | %s' % (archivo, e))