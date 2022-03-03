import sys, os
import datetime

from config_xlsx import PATH_LOG, CODM_XLSX, PROCESOS_GENERALES
from escribir_txt import salidaArchivoTxt, salidaLogTxt
from leerCoDmXLSX import LOG_PROCESO_CODM, leerArchivoCoDm
from validaciones_texto import validaFechaInput, encontrarDirectorio, encontrarArchivo, compruebaEncabezado, formatearFechaMesAnterior, sacarNombreArchivo, setearFechaInput

from crearXlsx import crearArchivoXlsx

def procesoCoDm(fechaInput, archivoXlsxInput, pathArchivoTxt, fechaInicioEntrada, fechaFinEntrada):

    hora = datetime.datetime.now()
    pathLogSalida = "CODM/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, 'CODM', fechaInput, hora.strftime("%Y%m%d%H%M"))
    try:
        dataTxt, encabezadoTxt, dataXlsx = leerArchivoCoDm(archivoXlsxInput, fechaInput, fechaInicioEntrada, fechaFinEntrada)
        formatoSalidaTxt = CODM_XLSX['SALIDA_TXT']
        logProceso = LOG_PROCESO_CODM
        
        salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoTxt, formatoSalidaTxt, fechaInput)
        
        if dataTxt:
            if salidaArchivoTxt(salidaTxt, dataTxt, encabezadoTxt):
                print("<a>&#128221;</a> Archivo {0} creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaTxt), len(dataTxt)))
                
                salidaXlsx = "{0}/{1}_{2}".format(pathArchivoTxt, fechaInput, CODM_XLSX['SALIDA_XLSX'])
                archivoProduccionXslx = ['PRODUCCION_{0}'.format(fechaInput), CODM_XLSX['ENCABEZADO_XLSX_PERIODO'], dataXlsx]
                if crearArchivoXlsx(salidaXlsx, [archivoProduccionXslx]):
                    print("<a>&#128221;</a> Archivo XLSX: {0}.xlsx creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaXlsx), len(dataXlsx)))

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo LOG: {0} Creado!".format(pathLogSalida))

    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):
    
    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [CODM_XLSX['ENCABEZADO_XLSX']]
    coordenadasEncabezado = [CODM_XLSX['COORDENADA_ENCABEZADO']]
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
    if len(sys.argv) == PROCESOS_GENERALES['CODM']['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[1])
        fechaRangoInicio = str(sys.argv[2])
        fechaRangoFin = str(sys.argv[3])
        archivoXlsCODM = str(sys.argv[4])
        pathArchivoTxt = str(sys.argv[5])
        setearFechaInput(fechaRangoInicio)
        setearFechaInput(fechaRangoFin)

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
        

        archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoXlsCODM])
        if archivosValidos and encabezadosValidos:
            procesoCoDm(fechaEntrada, archivoXlsCODM, pathArchivoTxt, fechaRangoInicio, fechaRangoFin)
        else:
            print('<a style="color:red">Error en Archivo:</a> {0}'.format(sacarNombreArchivo(archivoXlsCODM)))

    else:
        print("Error: El programa CODM necesita {0} parametros para su ejecucion".format(PROCESOS_GENERALES['CODM']['ARGUMENTOS_PROCESO']))

if __name__ == "__main__":
    main()