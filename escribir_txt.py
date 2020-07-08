from tqdm import tqdm
import csv

def archivo_txt(archivo_salida, data_txt, encabezado):
    try:
        with open(archivo_salida, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=';')
            writer.writerow(encabezado)
            for data in tqdm(iterable=data_txt, total = len(data_txt), desc='Escribiendo DATA', unit='Row'):
                writer.writerow(data)
        return True
    except Exception as e:
        print('Error al escribir archivo: %s | %s' % (archivo_salida, e))