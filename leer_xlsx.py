from openpyxl import load_workbook
from tqdm import tqdm

def archivo_fuga(archivo, periodo):
    try:
        data_txt = []
        encabezado = ['LPATTR_PER_RES', 'LLAVEA', 'LPATTR_COD_POLI', 'LPATTR_COD_ORIGEN', 'TIPO', 'LPATTR_COD_STAT', 'NRO_POLIZA', 'ESTADOPOLIZA', 'FECHAINICIOVIGENCIA', 'RUT_CRO', 'NOMBRE_CRO', 'FECHAPROCESO', 'CONSIDERAR_FUGA']
        encabezado_txt = ['CRR', 'FUGA', 'STOCK', 'RUT']
        # read_only=True para leer el archivo y consumir menos recursos
        # data_only=True para leer el valor de las celdas que tiene formulas
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        i = 0
        archivo_correcto = False
        data_txt
        for columna in hoja['A1:M1']:
            for celda in columna:
                dato = str(celda.value)
                if dato.upper() == encabezado[i]:
                    archivo_correcto = True
                else:
                    print('Error columna: %s' % i)
                    archivo_correcto = False
                i += 1
        if archivo_correcto:
            cantidad_columnas = len(tuple(hoja.rows))
            i = 1
            for columna in tqdm(iterable=hoja.rows, total = cantidad_columnas, desc='Leyendo DATA' , unit='Row'):
                data_xls = []
                control_data = False
                for celda in range(0, len(columna)):
                    considerar_fuga = str(columna[12].value)
                    if celda == 0:
                        data_xls.append(i)
                        periodo_data = str(columna[0].value)
                        if periodo == periodo_data:
                            control_data = True
                        else:
                            data_xls = []
                            control_data = False
                    if control_data and celda == 4:
                        tipo = str(columna[4].value)
                        # PREGUNTAR CUANDO CONSIDERAR_FUGA ESTE VACIO
                        if tipo.upper() == 'FUGA' and considerar_fuga.upper() != 'NO':
                            data_xls.append(1)
                            control_data = True
                        else:
                            control_data = False
                            data_xls = []
                    if control_data and celda == 5:
                        lpattr_cod_stat = str(columna[5].value)
                        # PREGUNTAR CUANDO CONSIDERAR_FUGA ESTE VACIO
                        if lpattr_cod_stat.upper() == 'NVIG' and considerar_fuga.upper() != 'NO':
                            data_xls.append(1)
                        else:
                            control_data = False
                            data_xls = []
                    if control_data and celda == 9:
                        rut_cro = str(columna[9].value)
                        if columna[9].value is not None:
                            rut_cro = rut_cro.partition("-")
                            numero_1, separador, numero_2 = rut_cro
                            rut = '%s%s' % (numero_1, numero_2) 
                            data_xls.append(rut)
                        else:
                            data_xls = []
                if len(data_xls) > 0:
                    data_txt.append(data_xls)
                    i += 1
        else:
            print('Error el archivo presenta incosistencias en el encabezado')
        return data_txt, encabezado_txt
    except Exception as e:
        print('Error al leer archivo: %s | %s' % (archivo, e))