import sys, os
import datetime

from config_xlsx import (CALIDAD_CONFIG_XLSX, CAMPANHAS_CONFIG_XLSX, FUGA_CONFIG_XLSX, GESTION_CONFIG_XLSX,
                         PATH_LOG, PATH_RAIZ, PROCESOS_GENERALES)
from escribir_txt import salidaArchivoTxt, salidaLogTxt

from leerDotacionXLSX import LOG_PROCESO_DOTACION, leerArchivoDotacion
from leerCalidadXSLX import LOG_PROCESO_CALIDAD, leerArchivoCalidad
from leerCampanhasEspecialesXLSX import (LOG_PROCESO_CAMPANHAS,
                                         leerArchivoCampanhasEsp)
from leerFugaXLSX import LOG_PROCESO_FUGA, leerArchivoFuga
from leerGestionXLSX import LOG_PROCESO_GESTION, leerArchivoGestion

from validaciones_texto import (compruebaEncabezado, encontrarArchivo,
                                encontrarDirectorio, validaFechaInput, setearFechaInput)


def procesoGenerico(fechaInput, archivoXlsxInput, pathArchivoTxt, procesoInput, *valoresExtraGestion):

    hora = datetime.datetime.now()
    pathLogSalida = "CRO/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, procesoInput, fechaInput, hora.strftime("H%HM%MS%S"))

    print("Iniciando Lectura del archivo de {0}...".format(archivoXlsxInput))
    try:
        if procesoInput == 'CALIDAD':
            dataXlsx, encabezadoXlsx = leerArchivoCalidad(archivoXlsxInput, fechaInput)
            formatoSalidaTxt = CALIDAD_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_CALIDAD
        elif procesoInput == 'CAMPAÑAS ESPEACIALES':
            dataXlsx, encabezadoXlsx = leerArchivoCampanhasEsp(archivoXlsxInput, fechaInput)
            formatoSalidaTxt = CAMPANHAS_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_CAMPANHAS
        elif procesoInput == 'FUGA':
            dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsxInput, fechaInput)
            formatoSalidaTxt = FUGA_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_FUGA
        elif procesoInput == 'GESTION':
            dataXlsx, encabezadoXlsx = leerArchivoGestion(archivoXlsxInput, fechaInput, valoresExtraGestion[0], valoresExtraGestion[1], valoresExtraGestion[2])
            formatoSalidaTxt = GESTION_CONFIG_XLSX['SALIDA_TXT']
            logProceso = LOG_PROCESO_GESTION
        
        salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaTxt, fechaInput)

        if dataXlsx:
            salidaArchivoTxt(salidaTxt, dataXlsx, encabezadoXlsx)

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo: {0}\{1} Creado!".format(PATH_RAIZ, pathLogSalida))
    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):

    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [CALIDAD_CONFIG_XLSX['ENCABEZADO_XLSX'], CAMPANHAS_CONFIG_XLSX['ENCABEZADO_XLSX'], FUGA_CONFIG_XLSX['ENCABEZADO_XLSX'], GESTION_CONFIG_XLSX['ENCABEZADO_XLSX'], GESTION_CONFIG_XLSX['ENCABEZADO_PROPIETARIOS_XLSX']]
    coordenadasEncabezado = [CALIDAD_CONFIG_XLSX['COORDENADA_ENCABEZADO'], CAMPANHAS_CONFIG_XLSX['COORDENADA_ENCABEZADO'], FUGA_CONFIG_XLSX['COORDENADA_ENCABEZADO'], GESTION_CONFIG_XLSX['COORDENADA_ENCABEZADO'], GESTION_CONFIG_XLSX['COORDENADA_ENCABEZADO_PROPIETARIO']]
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

def main():
    if len(sys.argv) == PROCESOS_GENERALES['CRO']['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[1])
        fechaRangoInicio = str(sys.argv[2])
        fechaRangoFin = str(sys.argv[3])
        archivoXlsCalidad = str(sys.argv[4])
        archivoXlsCampanas = str(sys.argv[5])
        archivoXlsFuga = str(sys.argv[6])
        archivoXlsGestion = str(sys.argv[7])
        archivoXlsPropietarios = str(sys.argv[8])
        pathArchivosTxt = str(sys.argv[9])
        procesosGenericos = ['CALIDAD', 'CAMPAÑAS ESPEACIALES', 'FUGA', 'GESTION']
        archivosProcesosGenericos = [archivoXlsCalidad, archivoXlsCampanas, archivoXlsFuga, archivoXlsGestion]

        if validaFechaInput(fechaEntrada) and setearFechaInput(fechaRangoInicio) and setearFechaInput(fechaRangoFin):
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

        archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoXlsCalidad, archivoXlsCampanas, archivoXlsFuga, archivoXlsGestion, archivoXlsPropietarios])

        if archivosValidos and encabezadosValidos:
            i = 0
            for proceso in procesosGenericos:
                procesoGenerico(fechaEntrada, archivosProcesosGenericos[i], pathArchivosTxt, proceso, fechaRangoInicio, fechaRangoFin, archivoXlsPropietarios)
                i +=1
        else:
            print("Error en Archivos de entrada")

    else:
        print("Error: El programa CRO necesita {0} parametros para su ejecucion".format(PROCESOS_GENERALES['CRO']['ARGUMENTOS_PROCESO']))

if __name__ == "__main__":
    main()
