import os.path
import sys

from leer_xlsx import archivo_fuga
from escribir_txt import archivo_txt
from validaciones_texto import validar_fecha

procesos = ['FUGA']
proceso = str(sys.argv[1])
if proceso.upper() in procesos:
    if len(sys.argv) == 3:
        fecha_x = str(sys.argv[2])
        fecha_a = fecha_x[0:4]
        fecha_b = fecha_x[4:6]
        if validar_fecha(fecha_a, fecha_b, fecha_x):
            archivo = '%s_FUGA_AGENCIA.xlsx' % fecha_x
            archivo_existe = os.path.isfile(archivo)
            salida_txt = 'FUGA%s.txt' % fecha_x
            if archivo_existe:
                print("Archivo: %s encontrado." % archivo)
                print("Iniciando Lectura...")
                data_txt, encabezado = archivo_fuga(archivo, fecha_x)
                if archivo_txt(salida_txt, data_txt, encabezado):
                    print("Archivo Creado !!")
                else:
                    print("Error!!")
            else:
                print('Error: Archivo %s no encontrado' % archivo)
        else:
            print("Error de fecha, formato correcto YYYYMM")
    else:
        print("Error: El programa "'"%s"'" necesita dos parametros" % proceso.upper())
else:
    print('Error: Proceso "'"%s"'" no encontrado' % proceso)










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