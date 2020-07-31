from openpyxl import load_workbook
from tqdm import tqdm

def validarFugaStock(correlativo, tipo, lpattrCodStat, considerarFuga, rut):
    rutNuevo = dict()
    rutNuevo["CRR"] = correlativo
    if tipo == 'FUGA' and  considerarFuga != 'NO':
        rutNuevo["FUGA"] = 1
    else:
        rutNuevo["FUGA"] = 0
    if  lpattrCodStat != 'NVIG' or considerarFuga == 'NO':
        rutNuevo["STOCK"] = 1
    else:
        rutNuevo["STOCK"] = 0
    rutNuevo["RUT"] = rut
    return rutNuevo

def existeFugaStock(tipo, lpattrCodStat, considerarFuga, rutExistente):
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
        i = 0
        archivo_correcto = False
        for fila in hoja['A1:M1']:
            for celda in fila:
                if str(celda.value).upper() == encabezadoXls[i]:
                    archivo_correcto = True
                else:
                    print('Error celda [M1]: %s ' % encabezadoXls[i])
                    archivo_correcto = False
                i += 1
        if archivo_correcto:
            i = 1
            filaSalidaXls = dict()
            for fila in tqdm(iterable=hoja.rows, total = len(tuple(hoja.rows)), desc='Leyendo DATA' , unit=' Fila'):
            # for fila in hoja.rows:
                if periodo == str(fila[0].value) and fila[9].value is not None:

                    rutMantisa, separador, dv = str(fila[9].value).partition("-")
                    rut = '%s%s' % (rutMantisa, dv)
                    tipo = str(fila[4].value).upper()
                    considerarFuga = str(fila[12].value).upper()
                    lpattrCodStat = str(fila[5].value).upper()

                    if filaSalidaXls.get(rut):
                        filaSalidaXls[rut] = existeFugaStock(tipo, lpattrCodStat, considerarFuga, filaSalidaXls[rut])
                    else:
                        filaSalidaXls[rut] = validarFugaStock(i, tipo, lpattrCodStat, considerarFuga, rut)
                        i += 1

            return filaSalidaXls, encabezadoTxt
        else:
            print('Error el archivo de FUGA presenta incosistencias en el encabezado')
            return False, False
    except Exception as e:
        print('Error al leer archivo: %s | %s' % (archivo, e))
