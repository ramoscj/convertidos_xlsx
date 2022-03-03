import sys, os
import datetime

from config_xlsx import PATH_LOG, PROACTIVA_CONFIG_XLSX, PROCESOS_GENERALES, PATH_RAIZ
from complementoCliente import COMPLEMENTO_CLIENTE_XLSX
from escribir_txt import salidaArchivoTxtProactiva, salidaLogTxt
from leerProactivaXLSX import LOG_PROCESO_PROACTIVA, leerArchivoProactiva
from validaciones_texto import validaFechaInput, encontrarDirectorio, encontrarArchivo, compruebaEncabezado, formatearFechaMesAnterior, sacarNombreArchivo

from crearXlsx import crearArchivoXlsx
from dataXlsxProactiva import dataXlsxReliquidacionesProactiva

def procesoProactiva(fechaInput, archivoXlsxInput, archivoComplementoCliente, pathArchivoTxt):

    hora = datetime.datetime.now()
    pathLogSalida = "PROACTIVA/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, 'PROACTIVA', fechaInput, hora.strftime("%Y%m%d%H%M"))
    try:
        dataProactivaTxt, encabezadoTxt, reliquidacionesTxt, encabezadoReliquidacionesTxt, dataArchivoXlsx = leerArchivoProactiva(archivoXlsxInput, fechaInput, archivoComplementoCliente)
        formatoSalidaTxt = PROACTIVA_CONFIG_XLSX['SALIDA_TXT']
        formatoSalidaReliquidaciones = PROACTIVA_CONFIG_XLSX['SALIDA_RELIQUIDACION']
        logProceso = LOG_PROCESO_PROACTIVA
        
        salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaTxt, fechaInput)
        salidaTxtReliquidaciones = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaReliquidaciones, fechaInput)
        salidaXlsx = "{0}/{1}_{2}".format(pathArchivoTxt, fechaInput, PROACTIVA_CONFIG_XLSX['SALIDA_XLSX'])
        mesAnterior = formatearFechaMesAnterior(fechaInput)
        
        if dataProactivaTxt:
            if salidaArchivoTxtProactiva(salidaTxt, dataProactivaTxt, encabezadoTxt) and salidaArchivoTxtProactiva(salidaTxtReliquidaciones, reliquidacionesTxt, encabezadoReliquidacionesTxt):
                print("<a>&#128221;</a> Archivo TXT: {0} creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaTxt), len(dataProactivaTxt)))
                print("<a>&#128221;</a> Archivo TXT: {0} creado con <strong> {1} reliquidaciones</strong>".format(sacarNombreArchivo(salidaTxtReliquidaciones), len(reliquidacionesTxt)))

            reliquidacionesXlsx = dataXlsxReliquidacionesProactiva(mesAnterior, reliquidacionesTxt)
            archivoProduccionPeriodo = ['PRODUCCION_{0}'.format(fechaInput), PROACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX_PERIODO'], dataArchivoXlsx ]
            archivoReliquidaciones = ['RELIQUIDACIONES_{0}'.format(mesAnterior.strftime("%Y%m")), PROACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX_REL'], reliquidacionesXlsx]
            if crearArchivoXlsx(salidaXlsx, [archivoProduccionPeriodo, archivoReliquidaciones]):
                print("<a>&#128221;</a> Archivo XLSX: {0}.xlsx creado con <strong> {1} registros y {2} reliquidaciones</strong>".format(sacarNombreArchivo(salidaXlsx), len(dataArchivoXlsx), len(reliquidacionesXlsx)))

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo LOG: {0} Creado!".format(pathLogSalida))

    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):
    
    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [PROACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX'], COMPLEMENTO_CLIENTE_XLSX['ENCABEZADO']]
    coordenadasEncabezado = [PROACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO'], COMPLEMENTO_CLIENTE_XLSX['COORDENADA_ENCABEZADO']]
    i = 0
    for archivo in archivosEntrada:
        if encontrarArchivo(archivo):
            print("<strong>Validando encabezado de {0}</strong>".format(sacarNombreArchivo(archivo)))
            archivoCorrecto = compruebaEncabezado(archivo, encabezadosArchivos[i], coordenadasEncabezado[i])

            if type(archivoCorrecto) is not dict:
                print("<a>&#9989;</a> Encabezado de Archivo: {0} OK!".format(sacarNombreArchivo(archivo)))
            else:
                encabezadosValidos = False
                for llave, valores in archivoCorrecto.items():
                    print('<a>&#10060;</a> {0}'.format(valores))
        else:
            print("Archivo: {0} NO Encontrado.".format(archivo))
            archivosValidos = False
        i += 1
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

        permisoPath = bool(os.access(pathArchivoTxt, os.W_OK))
        if not permisoPath:
            print('Error no tiene permisos de escritura en el directorio: {0}'.format(pathArchivoTxt))
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