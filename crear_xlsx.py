from openpyxl import load_workbook
from tqdm import tqdm
from tqdm.auto import tqdm
import os.path
import csv

def leer_xlsx(archivo):
    try:
        data_txt = []
        # read_only=True para leer el archivo y consumir menos recursos
        # data_only=True para leer el valor de las celdas que tiene formulas
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        cantidad_columnas = len(tuple(hoja.rows))
        for columna in tqdm(iterable=hoja.rows, total = cantidad_columnas, desc='Leyendo DATA' , unit='Row'):
            data_xls = []
            for fila in columna:
                data_xls.append(fila.value)
            data_txt.append(data_xls)
        return data_txt
    except Exception as e:
        raise Exception('Error al leer archivo: %s | %s' % (archivo, e))

def escribir_txt(archivo_salida, data_txt):
    try:
        with open(archivo_salida, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=';')
            for data in tqdm(iterable=data_txt, total = len(data_txt), desc='Escribiendo DATA', unit='Row'):
                writer.writerow(data)
        return True
    except Exception as e:
        raise Exception('Error al escribir archivo: %s | %s' % (archivo_salida, e))

def crear_txt():
    archivo = 'test.xlsx'
    archivo_existe = os.path.isfile(archivo)
    salida_txt = 'prueba.txt'
    print("Buscando archivo: %s" % archivo)

    if archivo_existe:
        print("Archivo: %s encontrado." % archivo)
        print("Iniciando Lectura...")
        data_txt = leer_xlsx(archivo)
        if escribir_txt(salida_txt, data_txt):
            print("Archivo Creado !!")
        else:
            print("Error!!")
        return True
    else:
        return False

crear_txt()










# with tqdm(total=len(my_list)) as pbar:
            #     for dato in data_txt:
            #         pbar.update(1)

# for i in tqdm(range(100001)):
#     print("", end='\r')

# loop = tqdm(total = 5000, position = 0, leave= False)
# for k in range(5000):
#     loop.set_description("Cargando...".format(k))
#     loop.update(1)
# loop.close()

# from tqdm import tqdm
# import requests

# chunk_size = 1024

# url = "http://www.tutorialspoint.com/python3/python_tutorial.pdf"

# req = requests.get(url, stream = True)

# total_size = int(req.headers['content-length'])

# with open("pythontutorial.pdf", "wb") as file:
#     for data in tqdm(iterable=req.iter_content(chunk_size=chunk_size), total = total_size/chunk_size, unit='KB'):
#         file.write(data)

# print("Download Completed !!!")