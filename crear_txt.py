import os.path
import sys

from leerFugaXLSX import leerArchivoFuga
from leerAsistenciaXLSX import leerArchivoAsistencia
from leerGestionXLSX import leerArchivoGestion
from leerCampanhasEspecialesXLSX import leerArchivoCampanhasEsp
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
            try:
                if procesoInput == 'FUGA':
                    dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsx, fechaInput)
                if procesoInput == 'ASISTENCIA':
                    dataXlsx, encabezadoXlsx = leerArchivoAsistencia(archivoXlsx, fechaInput)
                if procesoInput == 'CAMPANHA_ESPECIAL':
                    dataXlsx, encabezadoXlsx = leerArchivoCampanhasEsp(archivoXlsx, fechaInput)
                if salidaArchivoTxt(archivoTxtOutput, dataXlsx, encabezadoXlsx):
                    print("Archivo: %s Creado !!" % archivoTxtOutput)
            except Exception as e:
                print(e)
        else:
            print('Error: Archivo %s no encontrado' % archivoXlsx)

procesos = {'FUGA': {'argumentos': 2, 'archivoLecturaXls': '_FUGA_AGENCIA', 'archivoSalidaTxt': 'FUGA'}, 
            'ASISTENCIA': {'argumentos': 2, 'archivoLecturaXls': '_Asistencia_CRO', 'archivoSalidaTxt': 'ASISTENCIA'},
            'GESTION': {'argumentos': 4, 'archivoLecturaXls': 'Gestión CRO.xlsx', 'archivoSalidaTxt': 'GESTION'},
            'CAMPANHA_ESPECIAL': {'argumentos': 2, 'archivoLecturaXls': '_CampañasEspeciales_CRO', 'archivoSalidaTxt': 'PILOTO'}
            }
procesoInput = str(sys.argv[1]).upper()
if procesos.get(procesoInput):
    if len(sys.argv) == procesos[procesoInput]['argumentos'] + 1:
        if procesoInput == 'FUGA' or procesoInput == 'ASISTENCIA' or procesoInput == 'CAMPANHA_ESPECIAL':
            fechaEntrada = str(sys.argv[2])
            archivoXls = procesos[procesoInput]['archivoLecturaXls']
            archivoTxt = procesos[procesoInput]['archivoSalidaTxt']
            procesoGeneral(procesoInput, fechaEntrada, archivoXls, archivoTxt)
        elif procesoInput == 'GESTION':
            fechaEntrada = str(sys.argv[2])
            fechaRangoUno = str(sys.argv[3])
            fechaRangoDos = str(sys.argv[4])
            archivoXls = procesos[procesoInput]['archivoLecturaXls']
            if os.path.isfile(archivoXls):
                print("Archivo: %s encontrado." % archivoXls)
                print("Iniciando Lectura...")
                archivoTxt = ('%s%s.txt') % (procesos[procesoInput]['archivoSalidaTxt'], fechaEntrada)
                dataXlsx, encabezadoXlsx = leerArchivoGestion(archivoXls, fechaEntrada, fechaRangoUno, fechaRangoDos)
                if salidaArchivoTxt(archivoTxt, dataXlsx, encabezadoXlsx):
                    print("Archivo: GESTION Creado !!")
            else:
                print('Error: Archivo %s no encontrado' % archivoXls)
    else:
        print("Error: El programa "'"%s"'" necesita %s parametros para su ejecucion" % (procesoInput, procesos[procesoInput]['argumentos']))
else:
    print('Error: Proceso "'"%s"'" no encontrado' % procesoInput)