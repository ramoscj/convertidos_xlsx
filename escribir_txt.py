from tqdm import tqdm
import csv

def salidaArchivoTxt(ArchivoSalidaTxt, dataXlsx, encabezadoXlsx):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=';')
            writer.writerow(encabezadoXlsx)
            for pk, registro in tqdm(iterable=dataXlsx.items(), total = len(dataXlsx), desc='Escribiendo DATA', unit='Row'):
                writer.writerow(registro.values())
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))

def salidaArchivoTxtProactiva(ArchivoSalidaTxt, dataXlsx, encabezadoXlsx):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=';')
            writer.writerow(encabezadoXlsx)
            j = 1
            for pk, registro in tqdm(iterable=dataXlsx.items(), total = len(dataXlsx), desc='Escribiendo DATA', unit='Row'):
                data = [j]
                data += list(registro.values())
                writer.writerow(data)
                j += 1
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
