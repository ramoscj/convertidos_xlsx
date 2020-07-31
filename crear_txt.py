import os.path
import sys

from leerFugaXLSX import leerArchivoFuga
from leerAsistenciaXLSX import leerArchivoAsistencia
from escribir_txt import salidaArchivoTxt
from validaciones_texto import validaFechaInput

def procesoGeneral(procesoInput, fechaInput, archivoXlsxInput, archivoTxt):
    fechaYear = fechaInput[0:4]
    fechaMonth = fechaInput[4:6]
    if validaFechaInput(fechaYear, fechaMonth, fechaInput):
        archivoXlsx = '%s%s.xlsx' % (fechaInput, archivoXlsxInput)
        archivoTxtOutput = '%s%s.txt' % (archivoTxt, fechaInput)
        if os.path.isfile(archivoXlsx):
            print("Archivo: %s encontrado." % archivoXlsx)
            print("Iniciando Lectura...")
            if procesoInput == 'FUGA':
                dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsx, fechaInput)
            if procesoInput == 'ASISTENCIA':
                dataXlsx, encabezadoXlsx = leerArchivoAsistencia(archivoXlsx, fechaInput)
            if dataXlsx and salidaArchivoTxt(archivoTxtOutput, dataXlsx, encabezadoXlsx):
                print("Archivo Creado !!")
            else:
                print("Error al crear el archivo %s.txt" % archivoTxtOutput)
        else:
            print('Error: Archivo %s no encontrado' % archivoXlsx)




procesos = {'FUGA': {'argumentos': 3, 'archivoLecturaXls': '_FUGA_AGENCIA', 'archivoSalidaTxt': 'FUGA'}, 
            'ASISTENCIA': {'argumentos': 3, 'archivoLecturaXls': '_Asistencia_CRO', 'archivoSalidaTxt': 'ASISTENCIA'}
            }
procesoInput = str(sys.argv[1]).upper()
if procesos.get(procesoInput):
    if len(sys.argv) == procesos[procesoInput]['argumentos']:
        if procesoInput == 'FUGA' or procesoInput == 'ASISTENCIA':
            fechaEntrada = str(sys.argv[2])
            archivoXls = procesos[procesoInput]['archivoLecturaXls']
            archivoTxt = procesos[procesoInput]['archivoSalidaTxt']
            procesoGeneral(procesoInput, fechaEntrada, archivoXls, archivoTxt)
    else:
        print("Error: El programa "'"%s"'" necesita %s parametros para su ejecucion" % (procesoInput, procesos[procesoInput]['argumentos']))
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