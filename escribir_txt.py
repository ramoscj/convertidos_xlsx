from tqdm import tqdm
import csv

def salidaArchivoTxt(ArchivoSalidaTxt, dataXlsx, encabezadoXlsx):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=';')
            writer.writerow(encabezadoXlsx)
            for rut, x in tqdm(iterable=dataXlsx.items(), total = len(dataXlsx), desc='Escribiendo DATA', unit='Row'):
                writer.writerow(x.values())
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))

def salidaLogTxt(ArchivoSalidaTxt, dataXlsx):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter='\n')
            for rut, x in tqdm(iterable=dataXlsx.items(), total = len(dataXlsx), desc='Escribiendo LOG', unit='Row'):
                writer.writerow(x.values())
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))


import datetime
from dateutil.relativedelta import relativedelta
# from datetime import datetime

# datetime.today().replace(day=1)
primer = datetime.datetime(2020, 2, 5).replace(day=1).date()
ultimo = datetime.datetime(2020, 2, 5).replace(day=1).date()+relativedelta(months=1)+datetime.timedelta(days=-1)
# print(primer)
# print(ultimo)