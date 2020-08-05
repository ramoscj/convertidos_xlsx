from openpyxl import load_workbook
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx
from diccionariosDB import buscarRutEjecutivosDb
from config_xlsx import CAMPANHAS_CONFIG_XLSX

def leerArchivoCampanhasEsp(archivo, periodo):
    try:
        encabezadoXls = CAMPANHAS_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = CAMPANHAS_CONFIG_XLSX['ENCABEZADO_TXT']
        celda = CAMPANHAS_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
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
                if i >= 2 and fila[celda['RUT']].value is not None:
                    rut = formatearRut(str(fila[celda['RUT']].value))
                    if ejecutivosExistentesDb.get(rut):
                        numeroGestiones = fila[celda['NUMERO_GESTIONES']].value
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