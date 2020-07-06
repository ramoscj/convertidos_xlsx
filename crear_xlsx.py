from openpyxl import load_workbook
from tqdm import tqdm
from tqdm.auto import tqdm
import os.path
import csv

def crear_csv():
    archivo = 'test.xlsx'
    data_txt = []
    archivo_existe = os.path.isfile(archivo)
    salida_txt = 'prueba.txt'
    print("Buscando archivo: %s" % archivo)

    if archivo_existe:
        print("Archivo: %s encontrado." % archivo)
        print("Iniciando Lectura...")
        xls = load_workbook(archivo, read_only=True, data_only=True)
        nombre_hoja = xls.sheetnames
        hoja = xls[nombre_hoja[0]]
        cantidad_columnas = len(tuple(hoja.rows))
        for columna in tqdm(iterable=hoja.rows, total = cantidad_columnas, desc='Leyendo DATA ' , unit='Row'):
            data_xls = []
            for fila in columna:
                data_xls.append(fila.value)
            data_txt.append(data_xls)
        with open(salida_txt, 'w', newline='') as txt:
            writer = csv.writer(txt, delimiter=';')
            for data in tqdm(iterable=data_txt, total = len(data_txt), desc='Escribiendo DATA ', unit='Row'):
                writer.writerow(data)
        print("Archivo Creado !!")
        return True
    else:
        return False

crear_csv()










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