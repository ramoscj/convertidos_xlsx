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


# try:
#     if 1 == 1:
#         raise
#     else:
#         pass
#     print('sigo aqui')
# except Exception as e:
#     print('Error')