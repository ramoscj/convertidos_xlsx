from tqdm import tqdm
import csv

def salidaArchivoTxt(ArchivoSalidaTxt, dataXlsx, encabezadoXlsx):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='', encoding='UTF-8') as txt:
            writer = csv.writer(txt, delimiter=';')
            writer.writerow(encabezadoXlsx)
            for pk, registro in tqdm(iterable=dataXlsx.items(), total = len(dataXlsx), desc='Escribiendo DATA', unit='Row'):
            # for pk, registro in dataXlsx.items():
                writer.writerow(registro.values())
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))

def salidaArchivoTxtProactiva(ArchivoSalidaTxt, dataXlsx, encabezadoXlsx):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='', encoding='UTF-8') as txt:
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

def salidaLogTxt(ArchivoSalidaTxt, dataXlsx, encoding='UTF-8'):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter='\n')
            for rut, x in tqdm(iterable=dataXlsx.items(), total = len(dataXlsx), desc='Escribiendo LOG', unit='Row'):
            # for rut, x in dataXlsx.items():
                writer.writerow(x.values())
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))


def salidaInsertBulkCampanas(ArchivoSalidaTxt, dataXlsx, encabezado):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=',')
            writer.writerow(encabezado)
            for campana in tqdm(iterable=dataXlsx, total = len(dataXlsx), desc='Escribiendo ArchivoBulk', unit='Row'):
                writer.writerow(campana)
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))
    
    
def escribirArchivoTxt(ArchivoSalidaTxt, data, encabezado):
    try:
        with open(ArchivoSalidaTxt, 'w', newline='', encoding='UTF-8') as txt:
            writer = csv.writer(txt, delimiter=';')
            writer.writerow(encabezado)
            for registro in data:
                writer.writerow(registro)
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (ArchivoSalidaTxt, e))