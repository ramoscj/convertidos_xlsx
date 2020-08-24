import os.path
import sys

from leerFugaXLSX import leerArchivoFuga, LOG_PROCESO_FUGA
from leerAsistenciaXLSX import leerArchivoAsistencia, LOG_PROCESO_ASISTENCIA
from leerGestionXLSX import leerArchivoGestion, LOG_PROCESO_GESTION
from leerCampanhasEspecialesXLSX import leerArchivoCampanhasEsp, LOG_PROCESO_CAMPANHAS
from leerDotacionXLSX import leerArchivoDotacion, LOG_PROCESO_DOTACION

from escribir_txt import salidaArchivoTxt, salidaLogTxt
from validaciones_texto import validaFechaInput, formatearFechaMesSiguiente

from config_xlsx import PATH_XLSX, PATH_TXT, PATH_LOG

def procesoGeneral(procesoInput, fechaInput, archivoXlsxInput, archivoTxt):
    pathTxtSalida = PATH_TXT
    pathXlsxEntrada = PATH_XLSX
    
    if validaFechaInput(fechaInput):
        archivoXlsx = '%s%s%s.xlsx' % (pathXlsxEntrada, fechaInput, archivoXlsxInput)
        pathLogSalida = ('%slog_%s%s.txt') % (PATH_LOG, archivoTxt, fechaInput)
        if procesoInput == 'FUGA':
            archivoTxtOutput = '%s%s%s.txt' % (pathTxtSalida, archivoTxt, formatearFechaMesSiguiente(fechaInput))
        else:
            archivoTxtOutput = '%s%s%s.txt' % (pathTxtSalida, archivoTxt, fechaInput)
            
        if os.path.isfile(archivoXlsx):
            print("Archivo: %s encontrado." % archivoXlsx)
            print("Iniciando Lectura...")
            try:
                if procesoInput == 'FUGA':
                    dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsx, fechaInput)
                    logProceso = LOG_PROCESO_FUGA
                if procesoInput == 'ASISTENCIA':
                    dataXlsx, encabezadoXlsx = leerArchivoAsistencia(archivoXlsx, fechaInput)
                    dataXlsxDotacion, encabezadoXlsxDotacion = leerArchivoDotacion(fechaInput)
                    logProceso = LOG_PROCESO_ASISTENCIA
                    logProceso.update(LOG_PROCESO_DOTACION)
                if procesoInput == 'CAMPANHA_ESPECIAL':
                    dataXlsx, encabezadoXlsx = leerArchivoCampanhasEsp(archivoXlsx, fechaInput)
                    logProceso = LOG_PROCESO_CAMPANHAS
                if dataXlsx:
                    salidaArchivoTxt(archivoTxtOutput, dataXlsx, encabezadoXlsx)
                    if procesoInput == 'ASISTENCIA':
                        archivoTxtOutput = '%s%s%s.txt' % (pathTxtSalida, 'ICOM_CA_MTLFCC_', fechaInput)
                        salidaArchivoTxt(archivoTxtOutput, dataXlsxDotacion, encabezadoXlsxDotacion)
                    # LOG_PROCESO_GESTION.setdefault('SALIDA_TXT', {len(LOG_PROCESO_GESTION)+1: 'Archivo: %s creado!! ' % archivoTxtOutput})
                if salidaLogTxt(pathLogSalida, logProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)
            except Exception as e:
                print(e)
        else:
            print('Error: Archivo %s no encontrado' % archivoXlsx)

procesos = {'FUGA': {'argumentos': 2, 'archivoLecturaXls': '_FUGA_AGENCIA', 'archivoSalidaTxt': 'FUGA'}, 
            'ASISTENCIA': {'argumentos': 2, 'archivoLecturaXls': '_Asistencia_CRO', 'archivoSalidaTxt': 'ASISTENCIA'},
            'GESTION': {'argumentos': 2, 'archivoLecturaXls': 'Gestión CRO', 'archivoSalidaTxt': 'GESTION'},
            'CAMPANHA_ESPECIAL': {'argumentos': 2, 'archivoLecturaXls': '_CampañasEspeciales_CRO', 'archivoSalidaTxt': 'PILOTO'},
            'DOTACION': {'argumentos': 2, 'archivoLecturaXls': '_Asistencia_CRO', 'archivoSalidaTxt': 'ICOM_CA_CANAL_'}
            }
procesoInput = str(sys.argv[1]).upper()

if procesos.get(procesoInput):
    if len(sys.argv) == procesos[procesoInput]['argumentos'] + 1:
        if procesoInput == 'FUGA' or procesoInput == 'ASISTENCIA' or procesoInput == 'CAMPANHA_ESPECIAL' or procesoInput == 'DOTACION':
            fechaEntrada = str(sys.argv[2])
            archivoXls = procesos[procesoInput]['archivoLecturaXls']
            archivoTxt = procesos[procesoInput]['archivoSalidaTxt']
            procesoGeneral(procesoInput, fechaEntrada, archivoXls, archivoTxt)
        elif procesoInput == 'GESTION':
            fechaEntrada = str(sys.argv[2])
            # pathXlsxEntrada = 'test_xls/'
            pathXlsxEntrada = PATH_XLSX
            archivoXls = ('%s%s.xlsx') % (pathXlsxEntrada, procesos[procesoInput]['archivoLecturaXls'])
            if os.path.isfile(archivoXls):
                print("Archivo: %s encontrado." % archivoXls)
                print("Iniciando Lectura...")
                pathTxtSalida = PATH_TXT
                archivoTxt = ('%s%s%s.txt') % (pathTxtSalida, procesos[procesoInput]['archivoSalidaTxt'], fechaEntrada)
                pathLogSalida = ('%slog_%s_%s.txt') % (PATH_LOG, procesos[procesoInput]['archivoSalidaTxt'], fechaEntrada)
                dataXlsx, encabezadoXlsx = leerArchivoGestion(archivoXls, fechaEntrada)
                if dataXlsx and salidaArchivoTxt(archivoTxt, dataXlsx, encabezadoXlsx):
                    # print("Archivo: GESTION Creado !!")
                    LOG_PROCESO_GESTION.setdefault('SALIDA_TXT', {len(LOG_PROCESO_GESTION)+1: 'Archivo: %s creado!! ' % archivoTxt})
                erroresProceso = LOG_PROCESO_GESTION
                if salidaLogTxt(pathLogSalida, erroresProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)
            else:
                print('Error: Archivo %s no encontrado' % archivoXls)
    else:
        print("Error: El programa "'"%s"'" necesita %s parametros para su ejecucion" % (procesoInput, procesos[procesoInput]['argumentos']))
else:
    print('Error: Proceso "'"%s"'" no encontrado' % procesoInput)