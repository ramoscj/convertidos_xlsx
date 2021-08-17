import sys
import datetime

from config_xlsx import PATH_LOG, PROACTIVA_CONFIG_XLSX, PROCESOS_GENERALES, PATH_RAIZ
from complementoCliente import COMPLEMENTO_CLIENTE_XLSX
from escribir_txt import salidaArchivoTxtProactiva, salidaLogTxt
from leerProactivaXLSX import LOG_PROCESO_PROACTIVA, leerArchivoProactiva
from validaciones_texto import validaFechaInput, encontrarDirectorio, encontrarArchivo, compruebaEncabezado

def procesoProactiva(fechaInput, archivoXlsxInput, archivoComplementoCliente, pathArchivoTxt):

    hora = datetime.datetime.now()
    pathLogSalida = "PROACTIVA/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, 'PROACTIVA', fechaInput, hora.strftime("H%HM%MS%S"))
    print("Iniciando proceso PROACTIVA...")
    try:
        dataProactivaXlsx, encabezadoXlsx, reliquidacionesTxt, encabezadoReliquidacionesTxt = leerArchivoProactiva(archivoXlsxInput, fechaInput, archivoComplementoCliente)
        formatoSalidaTxt = PROACTIVA_CONFIG_XLSX['SALIDA_TXT']
        formatoSalidaReliquidaciones = PROACTIVA_CONFIG_XLSX['SALIDA_RELIQUIDACION']
        logProceso = LOG_PROCESO_PROACTIVA
        
        salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaTxt, fechaInput)
        salidaTxtReliquidaciones = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaReliquidaciones, fechaInput)

        if dataProactivaXlsx:
            salidaArchivoTxtProactiva(salidaTxt, dataProactivaXlsx, encabezadoXlsx)
            salidaArchivoTxtProactiva(salidaTxtReliquidaciones, reliquidacionesTxt, encabezadoReliquidacionesTxt)

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo: {0}\{1} Creado!".format(PATH_RAIZ, pathLogSalida))

    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):
    
    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [PROACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX'], COMPLEMENTO_CLIENTE_XLSX['ENCABEZADO']]
    coordenadasEncabezado = [PROACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO'], COMPLEMENTO_CLIENTE_XLSX['COORDENADA_ENCABEZADO']]
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
    if len(sys.argv) == PROCESOS_GENERALES['PROACTIVA']['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[1])
        archivoProactivaXls = str(sys.argv[2])
        archivoComplmentoCliente = str(sys.argv[3])
        pathArchivoTxt = str(sys.argv[4])

        if validaFechaInput(fechaEntrada):
            print("Fecha para el periodo %s OK!" % fechaEntrada)
        else:
            print("Fecha ingresada {0} incorrecta...".format(fechaEntrada))
            exit(1)

        salidaTxtDirectorio = encontrarDirectorio(pathArchivoTxt)
        if not salidaTxtDirectorio:
            print('Error Directorio: {0} no existe!'.format(str(pathArchivoTxt)))
            exit(1)
        
        archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoProactivaXls, archivoComplmentoCliente])

        if archivosValidos and encabezadosValidos:
            procesoProactiva(fechaEntrada, archivoProactivaXls, archivoComplmentoCliente, pathArchivoTxt)
        else:
            print("Error en Archivos de entrada")

    else:
        print("Error: El programa PROACTIVA necesita {0} parametros para su ejecucion".format(PROCESOS_GENERALES['PROACTIVA']['ARGUMENTOS_PROCESO']))

if __name__ == "__main__":
    main()