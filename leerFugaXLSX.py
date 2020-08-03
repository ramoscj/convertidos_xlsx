from openpyxl import load_workbook
from tqdm import tqdm

from validaciones_texto import formatearRut, validarEncabezadoXlsx

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
        encabezadoXls = ['LPATTR_PER_RES', 'LLAVEA', 'LPATTR_COD_POLI', 'LPATTR_COD_ORIGEN', 'TIPO', 'LPATTR_COD_STAT', 'NRO_POLIZA', 'ESTADOPOLIZA', 'FECHAINICIOVIGENCIA', 'RUT_CRO', 'NOMBRE_CRO', 'FECHAPROCESO', 'CONSIDERAR_FUGA']
        encabezadoTxt = ['CRR', 'FUGA', 'STOCK_PROXIMO_MES', 'RUT']
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        archivo_correcto = validarEncabezadoXlsx(hoja['A1:M1'], encabezadoXls)
        if archivo_correcto:
            i = 1
            filaSalidaXls = dict()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo FugaCRO' , unit=' Fila'):
            # for fila in hoja.rows:
                if periodo == str(fila[0].value) and fila[9].value is not None:

                    rut = rut = formatearRut(str(fila[9].value))
                    tipo = str(fila[4].value).upper()
                    considerarFuga = str(fila[12].value).upper()
                    lpattrCodStat = str(fila[5].value).upper()

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
