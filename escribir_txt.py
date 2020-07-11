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
        print('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))