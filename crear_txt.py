import os.path
import sys

from leerFugaXLSX import leerArchivoFuga, LOG_PROCESO_FUGA
from leerAsistenciaXLSX import leerArchivoAsistencia, LOG_PROCESO_ASISTENCIA
from leerGestionXLSX import leerArchivoGestion, LOG_PROCESO_GESTION
from leerCampanhasEspecialesXLSX import leerArchivoCampanhasEsp, LOG_PROCESO_CAMPANHAS
from leerDotacionXLSX import leerArchivoDotacion, LOG_PROCESO_DOTACION
from leerCalidadXSLX import leerArchivoCalidad, LOG_PROCESO_CALIDAD

from escribir_txt import salidaArchivoTxt, salidaLogTxt
from validaciones_texto import validaFechaInput, formatearFechaMesSiguiente

from config_xlsx import FUGA_CONFIG_XLSX, ASISTENCIA_CONFIG_XLSX, GESTION_CONFIG_XLSX, CAMPANHAS_CONFIG_XLSX, CALIDAD_CONFIG_XLSX, DOTACION_CONFIG_XLSX
from config_xlsx import PATH_XLSX, PATH_TXT, PATH_LOG

def procesoGeneral(procesoInput, fechaInput, archivoXlsxInput, archivoTxt):
    pathTxtSalida = PATH_TXT
    pathXlsxEntrada = PATH_XLSX

    if validaFechaInput(fechaInput):
        archivoXlsx = '%s%s%s.xlsx' % (pathXlsxEntrada, fechaInput, archivoXlsxInput)
        pathLogSalida = ('%slog_%s%s.txt') % (PATH_LOG, archivoTxt, fechaInput)
        archivoTxtOutput = '%s%s%s.txt' % (pathTxtSalida, archivoTxt, fechaInput)

        if os.path.isfile(archivoXlsx):
            print("Archivo: %s encontrado." % archivoXlsx)
            print("Iniciando Lectura...")
            try:
                if procesoInput == 'FUGA':
                    mesSiguiente = formatearFechaMesSiguiente(fechaInput)
                    archivoTxtOutput = '%s%s%s.txt' % (pathTxtSalida, archivoTxt, mesSiguiente.strftime("%Y%m"))
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
                if procesoInput == 'CALIDAD':
                    dataXlsx, encabezadoXlsx = leerArchivoCalidad(archivoXlsx, fechaInput)
                    logProceso = LOG_PROCESO_CALIDAD

                if dataXlsx:
                    salidaArchivoTxt(archivoTxtOutput, dataXlsx, encabezadoXlsx)
                    if procesoInput == 'ASISTENCIA':
                        archivoTxtOutput = '%s%s%s.txt' % (pathTxtSalida, DOTACION_CONFIG_XLSX['SALIDA_TXT'], fechaInput)
                        salidaArchivoTxt(archivoTxtOutput, dataXlsxDotacion, encabezadoXlsxDotacion)

                if salidaLogTxt(pathLogSalida, logProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)
            except Exception as e:
                print(e)
        else:
            print('Error: Archivo %s no encontrado' % archivoXlsx)

procesos = {'FUGA': FUGA_CONFIG_XLSX,
            'ASISTENCIA': ASISTENCIA_CONFIG_XLSX,
            'GESTION': GESTION_CONFIG_XLSX,
            'CAMPANHA_ESPECIAL': CAMPANHAS_CONFIG_XLSX,
            'CALIDAD': CALIDAD_CONFIG_XLSX
            }
procesoInput = str(sys.argv[1]).upper()

if procesos.get(procesoInput):
    if len(sys.argv) == procesos[procesoInput]['ARGUMENTOS_PROCESO'] + 1:
        if procesoInput == 'FUGA' or procesoInput == 'ASISTENCIA' or procesoInput == 'CAMPANHA_ESPECIAL' or procesoInput == 'CALIDAD':
            fechaEntrada = str(sys.argv[2])
            archivoXls = procesos[procesoInput]['ENTRADA_XLSX']
            archivoTxt = procesos[procesoInput]['SALIDA_TXT']
            procesoGeneral(procesoInput, fechaEntrada, archivoXls, archivoTxt)
        elif procesoInput == 'GESTION':
            fechaEntrada = str(sys.argv[2])
            fechaRangoUno = str(sys.argv[3])
            fechaRangoDos = str(sys.argv[4])
            pathXlsxEntrada = PATH_XLSX
            archivoXls = ('%s%s.xlsx') % (pathXlsxEntrada, procesos[procesoInput]['ENTRADA_XLSX'])
            if os.path.isfile(archivoXls):
                print("Archivo: %s encontrado." % archivoXls)
                print("Iniciando Lectura...")
                pathTxtSalida = PATH_TXT
                archivoTxt = ('%s%s%s.txt') % (pathTxtSalida, procesos[procesoInput]['SALIDA_TXT'], fechaEntrada)
                pathLogSalida = ('%slog_%s_%s.txt') % (PATH_LOG, procesos[procesoInput]['SALIDA_TXT'], fechaEntrada)
                dataXlsx, encabezadoXlsx = leerArchivoGestion(archivoXls, fechaEntrada, fechaRangoUno, fechaRangoDos)
                if dataXlsx and salidaArchivoTxt(archivoTxt, dataXlsx, encabezadoXlsx):
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