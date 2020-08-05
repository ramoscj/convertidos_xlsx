from openpyxl import load_workbook
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx
from config_xlsx import FUGA_CONFIG_XLSX, PATH_XLSX

def validarFugaStock(correlativo, tipo, lpattrCodStat, considerarFuga, rut):
    rutNuevo = dict()
    rutNuevo["CRR"] = correlativo
    if tipo == 'FUGA' and considerarFuga != 'NO':
        rutNuevo["FUGA"] = 1
    else:
        rutNuevo["FUGA"] = 0
    if  lpattrCodStat != 'NVIG' or considerarFuga == 'NO':
        rutNuevo["STOCK"] = 1
    else:
        rutNuevo["STOCK"] = 0
    rutNuevo["RUT"] = rut
    return rutNuevo

def existeRut(tipo, lpattrCodStat, considerarFuga, rutExistente):
    if tipo == 'FUGA' and  considerarFuga != 'NO':
        rutExistente["FUGA"] += 1
    elif  lpattrCodStat != 'NVIG' or considerarFuga == 'NO':
        rutExistente["STOCK"] += 1
    return rutExistente

def leerArchivoFuga(archivo, periodo):
    try:
        encabezadoXls = FUGA_CONFIG_XLSX['ENCABEZADO_XLSX']
        encabezadoTxt = FUGA_CONFIG_XLSX['ENCABEZADO_TXT']
        columna = FUGA_CONFIG_XLSX['COLUMNAS_PROCESO_XLSX']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        archivo_correcto = validarEncabezadoXlsx(hoja['A1:M1'], encabezadoXls)
        if archivo_correcto:
            i = 1
            filaSalidaXls = dict()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo FugaCRO' , unit=' Fila'):
            # for fila in hoja.rows:
                if periodo == str(fila[columna['LPATTR_PER_RES']].value) and fila[columna['RUT_CRO']].value is not None:

                    rut = formatearRut(str(fila[columna['RUT_CRO']].value))
                    tipo = str(fila[columna['TIPO']].value).upper()
                    considerarFuga = str(fila[columna['CONSIDERAR_FUGA']].value).upper()
                    lpattrCodStat = str(fila[columna['LPATTR_COD_STAT']].value).upper()

                    if filaSalidaXls.get(rut):
                        filaSalidaXls[rut] = existeRut(tipo, lpattrCodStat, considerarFuga, filaSalidaXls[rut])
                    else:
                        filaSalidaXls[rut] = validarFugaStock(i, tipo, lpattrCodStat, considerarFuga, rut)
                        i += 1

            return filaSalidaXls, encabezadoTxt
        else:
            raise Exception('Incosistencias en el encabezado')
    except Exception as e:
        raise Exception('Error al leer archivo: %s | %s' % (archivo, e))
