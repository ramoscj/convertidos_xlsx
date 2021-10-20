import sys, os
import datetime

from config_xlsx import (CALIDAD_CONFIG_XLSX, CAMPANHAS_CONFIG_XLSX, FUGA_CONFIG_XLSX, GESTION_CONFIG_XLSX,
                         PATH_LOG, PATH_RAIZ, PROCESOS_GENERALES, CODM_XLSX, CAMPANHAS_PRIORITARIAS)
from escribir_txt import salidaArchivoTxt, salidaLogTxt

from leerDotacionXLSX import LOG_PROCESO_DOTACION, leerArchivoDotacion
from leerCalidadXSLX import LOG_PROCESO_CALIDAD, leerArchivoCalidad
from leerCampanhasEspecialesXLSX import (LOG_PROCESO_CAMPANHAS,
                                         leerArchivoCampanhasEsp)
from leerFugaXLSX import LOG_PROCESO_FUGA, leerArchivoFuga
from leerGestionXLSX import LOG_PROCESO_GESTION, leerArchivoGestion
from leerCoDmXLSX import LOG_PROCESO_CODM, leerArchivoCoDm
from leerCampanasPrioritariasXLSX import LOG_PROCESO_PRIORITARIAS, leerArchivoPrioritarias

from validaciones_texto import (compruebaEncabezado, encontrarArchivo,
                                encontrarDirectorio, validaFechaInput, setearFechaInput, formatearFechaMesAnterior)

procesos = {

    "FUGA": {
                'PROCESO': PROCESOS_GENERALES['FUGA']['ARGUMENTOS_PROCESO'],
                'ENCABEZADO': FUGA_CONFIG_XLSX['ENCABEZADO_XLSX'],
                'COORDENADA_ENCABEZADO': FUGA_CONFIG_XLSX['COORDENADA_ENCABEZADO'],
                },
    "GESTION": {
                'PROCESO': PROCESOS_GENERALES['GESTION']['ARGUMENTOS_PROCESO'],
                'ENCABEZADO': GESTION_CONFIG_XLSX['ENCABEZADO_XLSX'],
                'COORDENADA_ENCABEZADO': GESTION_CONFIG_XLSX['COORDENADA_ENCABEZADO'],
                'ENCABEZADO_PROPIETARIOS' :GESTION_CONFIG_XLSX['ENCABEZADO_PROPIETARIOS_XLSX'],
                'COORDENADA_PROPIETARIOS' :GESTION_CONFIG_XLSX['COORDENADA_ENCABEZADO_PROPIETARIO'],
                },
    "CAMPANHA_ESPECIAL": {
                'PROCESO': PROCESOS_GENERALES['CAMPANHA_ESPECIAL']['ARGUMENTOS_PROCESO'],
                'ENCABEZADO': CAMPANHAS_CONFIG_XLSX['ENCABEZADO_XLSX'],
                'COORDENADA_ENCABEZADO': CAMPANHAS_CONFIG_XLSX['COORDENADA_ENCABEZADO'],
                },
    "CALIDAD": {
                'PROCESO': PROCESOS_GENERALES['CALIDAD']['ARGUMENTOS_PROCESO'],
                'ENCABEZADO': CALIDAD_CONFIG_XLSX['ENCABEZADO_XLSX'],
                'COORDENADA_ENCABEZADO': CALIDAD_CONFIG_XLSX['COORDENADA_ENCABEZADO'],
                },
    "CAMPANHA_PRIORITARIA": {
                'PROCESO': PROCESOS_GENERALES['CAMPANHA_PRIORITARIA']['ARGUMENTOS_PROCESO'],
                'ENCABEZADO': CAMPANHAS_PRIORITARIAS['ENCABEZADO_XLSX'],
                'COORDENADA_ENCABEZADO': CAMPANHAS_PRIORITARIAS['COORDENADA_ENCABEZADO'],
                },
    "CODM":     {
                'PROCESO': PROCESOS_GENERALES['CODM']['ARGUMENTOS_PROCESO'],
                'ENCABEZADO': CODM_XLSX['ENCABEZADO_XLSX'],
                'COORDENADA_ENCABEZADO': CODM_XLSX['COORDENADA_ENCABEZADO'],
                },
        
}

def procesoGenerico(fechaInput, archivoXlsxInput, pathArchivoTxt, procesoInput, *valoresExtraGestion):

    hora = datetime.datetime.now()
    pathLogSalida = "CRO/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, procesoInput, fechaInput, hora.strftime("%Y%m%d%H%M"))

    print("Iniciando Lectura del archivo de {0}...".format(archivoXlsxInput))
    try:
        if procesoInput == 'CALIDAD':
            dataXlsx, encabezadoXlsx = leerArchivoCalidad(archivoXlsxInput, fechaInput)
            formatoSalidaTxt = CALIDAD_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_CALIDAD
        elif procesoInput == 'CAMPANHA_ESPECIAL':
            dataXlsx, encabezadoXlsx = leerArchivoCampanhasEsp(archivoXlsxInput, fechaInput)
            formatoSalidaTxt = CAMPANHAS_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_CAMPANHAS
        elif procesoInput == 'FUGA':
            fechaMesAnterior = formatearFechaMesAnterior(fechaInput)
            dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsxInput, fechaMesAnterior.strftime("%Y%m"))
            formatoSalidaTxt = FUGA_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_FUGA
        elif procesoInput == 'CODM':
            dataXlsx, encabezadoXlsx = leerArchivoCoDm(archivoXlsxInput, fechaInput, valoresExtraGestion[0], valoresExtraGestion[1])
            formatoSalidaTxt = CODM_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_CODM
        elif procesoInput == 'GESTION':
            dataXlsx, encabezadoXlsx = leerArchivoGestion(archivoXlsxInput, fechaInput, valoresExtraGestion[0], valoresExtraGestion[1], valoresExtraGestion[2])
            formatoSalidaTxt = GESTION_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_GESTION
        elif procesoInput == 'CAMPANHA_PRIORITARIA':
            dataXlsx, encabezadoXlsx = leerArchivoPrioritarias(archivoXlsxInput)
            formatoSalidaTxt = CAMPANHAS_PRIORITARIAS['SALIDA_TXT']
            logProceso = LOG_PROCESO_PRIORITARIAS

        
        salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaTxt, fechaInput)

        if dataXlsx:
            salidaArchivoTxt(salidaTxt, dataXlsx, encabezadoXlsx)
        elif procesoInput == 'CAMPANHA_PRIORITARIA':
            salidaArchivoTxt(salidaTxt, dataXlsx, encabezadoXlsx)

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo: {0}\{1} Creado!".format(PATH_RAIZ, pathLogSalida))
    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: [], encabezadosArchivos: [], coordenadasEncabezado: []):

    archivosValidos = True
    encabezadosValidos = True
    i = 0
    print("-----------------------------------------------------")
    for archivo in archivosEntrada:
        if encontrarArchivo(archivo):
            print("Archivo: {0} Encontrado!".format(archivo))
            archivoCorrecto = compruebaEncabezado(archivo, encabezadosArchivos[i], coordenadasEncabezado[i])

            if type(archivoCorrecto) is not dict:
                print(".- Encabezado de Archivo: {0} OK!".format(archivo))
            else:
                encabezadosValidos = False
                for llave, valores in archivoCorrecto.items():
                    print('.- {0}'.format(valores))
        else:
            print("Archivo: {0} NO Encontrado.".format(archivo))
            archivosValidos = False
        i += 1
    print("-----------------------------------------------------")
    return archivosValidos, encabezadosValidos

def main(procesoInput):
    if len(sys.argv) == PROCESOS_GENERALES[procesoInput]['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[2])
        proceso = 0

        if procesoInput == 'CALIDAD' or procesoInput == 'CAMPANHA_ESPECIAL' or procesoInput == 'FUGA' or procesoInput == 'CAMPANHA_PRIORITARIA':
            archivoXlsEntrada = str(sys.argv[3])
            pathArchivosTxt = str(sys.argv[4])
            proceso = 1
        elif procesoInput == 'CODM':
            fechaRangoInicio = str(sys.argv[3])
            fechaRangoFin = str(sys.argv[4])
            archivoXlsCODM = str(sys.argv[5])
            pathArchivosTxt = str(sys.argv[6])
            setearFechaInput(fechaRangoInicio)
            setearFechaInput(fechaRangoFin)
            proceso = 2
        elif procesoInput == 'GESTION':
            fechaRangoInicio = str(sys.argv[3])
            fechaRangoFin = str(sys.argv[4])
            archivoXlsGestion = str(sys.argv[5])
            archivoXlsPropietarios = str(sys.argv[6])
            pathArchivosTxt = str(sys.argv[7])
            setearFechaInput(fechaRangoInicio)
            setearFechaInput(fechaRangoFin)
            proceso = 3

        if validaFechaInput(fechaEntrada):
            print("Fecha para el periodo %s OK!" % fechaEntrada)
        else:
            print("Fecha ingresada {0} incorrecta...".format(fechaEntrada))
            exit(1)
            
        salidaTxtDirectorio = encontrarDirectorio(pathArchivosTxt)
        if not salidaTxtDirectorio:
            print('Error Directorio: {0} no existe!'.format(str(pathArchivosTxt)))
            exit(1)

        permisoPath = bool(os.access(pathArchivosTxt, os.W_OK))
        if not permisoPath:
            print('Error no tiene permisos de escritura en el directorio: {0}'.format(pathArchivosTxt))
            exit(1)

        if proceso == 1:
            archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoXlsEntrada], [procesos[procesoInput]['ENCABEZADO']], [procesos[procesoInput]['COORDENADA_ENCABEZADO']])
            if archivosValidos and encabezadosValidos:
                procesoGenerico(fechaEntrada, archivoXlsEntrada, pathArchivosTxt, procesoInput)
            else:
                print("Error en Archivo: {0}".format(archivoXlsEntrada))
        elif proceso == 2:
            archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoXlsCODM], [procesos[procesoInput]['ENCABEZADO']], [procesos[procesoInput]['COORDENADA_ENCABEZADO']])
            if archivosValidos and encabezadosValidos:
                procesoGenerico(fechaEntrada, archivoXlsCODM, pathArchivosTxt, procesoInput, fechaRangoInicio, fechaRangoFin)
            else:
                print("Error en Archivo: {0}".format(archivoXlsEntrada))
        elif proceso == 3:
            archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoXlsGestion, archivoXlsPropietarios], [procesos[procesoInput]['ENCABEZADO'], procesos[procesoInput]['ENCABEZADO_PROPIETARIOS']], [procesos[procesoInput]['COORDENADA_ENCABEZADO'], procesos[procesoInput]['COORDENADA_PROPIETARIOS']])
            if archivosValidos and encabezadosValidos:
                procesoGenerico(fechaEntrada, archivoXlsGestion, pathArchivosTxt, procesoInput, fechaRangoInicio, fechaRangoFin, archivoXlsPropietarios)
            else:
                print("Error en Archivo: {0}".format(archivoXlsEntrada))

    else:
        print("Error: El programa {0} necesita {1} parametros para su ejecucion".format(procesoInput, PROCESOS_GENERALES[procesoInput]['ARGUMENTOS_PROCESO']))

if __name__ == "__main__":
    procesoInput = str(sys.argv[1]).upper()
    if procesos.get(procesoInput):
        main(procesoInput)
    else:
        print('Error: Proceso "' "{0}" '" no encontrado'.format(procesoInput))
