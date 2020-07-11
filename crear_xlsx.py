import os.path
import sys

from leer_xlsx import leerArchivoFuga
from escribir_txt import salidaArchivoTxt
from validaciones_texto import validaFechaInput

procesos = ['FUGA']
procesoInput = str(sys.argv[1])
if procesoInput.upper() in procesos:
    if len(sys.argv) == 3:
        fechaInput = str(sys.argv[2])
        fechaYear = fechaInput[0:4]
        fechaMonth = fechaInput[4:6]
        if validaFechaInput(fechaYear, fechaMonth, fechaInput):
            archivoXlsx = '%s_FUGA_AGENCIA.xlsx' % fechaInput
            archivoTxtxOutput = 'FUGA%s.txt' % fechaInput
            if os.path.isfile(archivoXlsx):
                print("Archivo: %s encontrado." % archivoXlsx)
                print("Iniciando Lectura...")
                dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsx, fechaInput)
                if dataXlsx and salidaArchivoTxt(archivoTxtxOutput, dataXlsx, encabezadoXlsx):
                    print("Archivo Creado !!")
                else:
                    print("Error!!")
            else:
                print('Error: Archivo %s no encontrado' % archivoXlsx)
        else:
            print("Error de fecha, formato correcto YYYYMM")
    else:
        print("Error: El programa "'"%s"'" necesita dos parametros" % procesoInput.upper())
else:
    print('Error: Proceso "'"%s"'" no encontrado' % procesoInput)










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