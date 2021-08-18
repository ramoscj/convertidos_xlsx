import sys
import datetime

from config_xlsx import PATH_LOG, REACTIVA_CONFIG_XLSX, PROCESOS_GENERALES, PATH_RAIZ
from complementoCliente import COMPLEMENTO_CLIENTE_XLSX
from escribir_txt import salidaArchivoTxtProactiva, salidaLogTxt
from leerReactivaXLSX import LOG_PROCESO_REACTIVA, leerArchivoReactiva
from validaciones_texto import validaFechaInput, encontrarDirectorio, encontrarArchivo, compruebaEncabezado

def procesoReactiva(fechaInput, fechaRangoUno, fechaRangoDos, archivoXlsxInput, archivoCertificacionXls, archivoComplementoCliente, pathArchivoTxt):
    
    hora = datetime.datetime.now()
    pathLogSalida = "REACTIVA/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, 'REACTIVA', fechaInput, hora.strftime("H%HM%MS%S"))
    print("Iniciando proceso REACTIVA...")
    try:
        dataReactivaTxt = leerArchivoReactiva(archivoXlsxInput, fechaInput, fechaRangoUno, fechaRangoDos, archivoCertificacionXls, archivoComplementoCliente)
        logProceso = LOG_PROCESO_REACTIVA
        
        if dataReactivaTxt:
            for data in dataReactivaTxt:
                salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoTxt, data['NOMBRE_ARCHIVO'], fechaInput)
                salidaArchivoTxtProactiva(salidaTxt, data['DATA'], data['ENCABEZADO'])

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo: {0}\{1} Creado!".format(PATH_RAIZ, pathLogSalida))

    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):
    
    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [REACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX'], COMPLEMENTO_CLIENTE_XLSX['ENCABEZADO'], REACTIVA_CONFIG_XLSX['ARCHIVO_BASE_CERTIFICACION']['ENCABEZADO']]
    coordenadasEncabezado = [REACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO'], COMPLEMENTO_CLIENTE_XLSX['COORDENADA_ENCABEZADO'], REACTIVA_CONFIG_XLSX['ARCHIVO_BASE_CERTIFICACION']['COORDENADA_ENCABEZADO']]
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
    if len(sys.argv) == PROCESOS_GENERALES['REACTIVA']['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[1])
        fechaRangoUno = str(sys.argv[2])
        fechaRangoDos = str(sys.argv[3])
        archivoReactivaXls = str(sys.argv[4])
        archivoCertificacionXls = str(sys.argv[5])
        archivoComplementoXls = str(sys.argv[6])
        pathArchivoTxt = str(sys.argv[7])

        if validaFechaInput(fechaEntrada):
            print("Fecha para el periodo %s OK!" % fechaEntrada)
        else:
            print("Fecha ingresada {0} incorrecta...".format(fechaEntrada))
            exit(1)

        salidaTxtDirectorio = encontrarDirectorio(pathArchivoTxt)
        if not salidaTxtDirectorio:
            print('Error Directorio: {0} no existe!'.format(str(pathArchivoTxt)))
            exit(1)

        archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoReactivaXls, archivoComplementoXls, archivoCertificacionXls])

        if archivosValidos and encabezadosValidos:
            procesoReactiva(fechaEntrada, fechaRangoUno, fechaRangoDos, archivoReactivaXls, archivoCertificacionXls, archivoComplementoXls, pathArchivoTxt)
        else:
            print("Error en Archivos de entrada")
    else:
        print("Error: El programa REACTIVA necesita {0} parametros para su ejecucion".format(PROCESOS_GENERALES['REACTIVA']['ARGUMENTOS_PROCESO']))

if __name__ == "__main__":
    main()