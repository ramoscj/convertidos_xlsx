import sys, os
import datetime

from config_xlsx import (ASISTENCIA_CONFIG_XLSX, DOTACION_CONFIG_XLSX,
                         PATH_LOG, PATH_RAIZ, PROCESOS_GENERALES)
from escribir_txt import salidaArchivoTxt, salidaLogTxt
from leerAsistenciaXLSX import LOG_PROCESO_ASISTENCIA, leerArchivoAsistencia
from leerDotacionXLSX import LOG_PROCESO_DOTACION, leerArchivoDotacion
from validaciones_texto import (compruebaEncabezado, encontrarArchivo,
                                encontrarDirectorio, validaFechaInput, sacarNombreArchivo)


hora = datetime.datetime.now()

def procesoAsistencia(fechaInput, archivoXlsxInput, pathArchivoTxt):
    
    procesoInput = 'ASISTENCIA'
    pathLogSalida = "CRO/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, procesoInput, fechaInput, hora.strftime("%Y%m%d%H%M"))

    print("<strong>Iniciando Lectura del archivo {0}</strong>".format(sacarNombreArchivo(archivoXlsxInput)))
    try:
        dataTxt, encabezadoXlsx = leerArchivoAsistencia(archivoXlsxInput, fechaInput)
        salidaTxtAsistencia = "{0}/{1}{2}.txt".format(pathArchivoTxt, ASISTENCIA_CONFIG_XLSX['SALIDA_TXT'], fechaInput)
        logProceso = LOG_PROCESO_ASISTENCIA

        if dataTxt:
            salidaArchivoTxt(salidaTxtAsistencia, dataTxt, encabezadoXlsx)
            print("<a>&#128221;</a> Archivo TXT: {0} creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaTxtAsistencia), len(dataTxt)))

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo: {0} Creado!".format(pathLogSalida))
            print("-----------------------------------------------------")
            
    except Exception as e:
        print(e)

def procesoDotacion(fechaInput, pathArchivoTxt):

    procesoInput = 'DOTACION'
    pathLogSalidaDotacion = "CRO/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, procesoInput, fechaInput, hora.strftime("%Y%m%d%H%M"))
    salidaTxtDotacion = "{0}/{1}{2}.txt".format(pathArchivoTxt, DOTACION_CONFIG_XLSX['SALIDA_TXT'], fechaInput)

    print("<strong>Iniciando proceso de Archivo DOTACION</strong>")
    dataTxt, encabezadoXlsxDotacion = leerArchivoDotacion(fechaInput)
    try:
        if dataTxt:
            salidaArchivoTxt(salidaTxtDotacion, dataTxt, encabezadoXlsxDotacion)
            print("<a>&#128221;</a> Archivo TXT: {0} creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaTxtDotacion), len(dataTxt)))
            
        logProceso = LOG_PROCESO_DOTACION
        if salidaLogTxt(pathLogSalidaDotacion, logProceso):
            print("Archivo: {0} Creado!".format(pathLogSalidaDotacion))
    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):
    
    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [ASISTENCIA_CONFIG_XLSX['ENCABEZADO_XLSX']]
    coordenadasEncabezado = [ASISTENCIA_CONFIG_XLSX['COORDENADA_ENCABEZADO']]
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
    if len(sys.argv) == PROCESOS_GENERALES['DOTACION']['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[1])
        archivoXlsAsistencia = str(sys.argv[2])
        pathArchivoTxtAsistencia = str(sys.argv[3])
        pathArchivoTxtDotacion = str(sys.argv[4])
        pathNoEncontrado = []
        pathNoPermisos = []
        directorioNumero = []

        if validaFechaInput(fechaEntrada):
            print("Fecha para el periodo %s OK!" % fechaEntrada)
        else:
            print("Fecha ingresada {0} incorrecta...".format(fechaEntrada))
            exit(1)
        
        salidaTxtAsistencia = encontrarDirectorio(pathArchivoTxtAsistencia)
        salidaTxtDotacion = encontrarDirectorio(pathArchivoTxtDotacion)
        if not salidaTxtAsistencia or not salidaTxtDotacion:
            if not salidaTxtAsistencia:
                pathNoEncontrado.append(pathArchivoTxtAsistencia)
                directorioNumero.append(1)
            if not salidaTxtDotacion:
                pathNoEncontrado.append(pathArchivoTxtDotacion)
                directorioNumero.append(2)
            i = 0
            for path in pathNoEncontrado:
                print('Error en el Directorio {0}: {1} no existe!'.format(directorioNumero[i], str(path)))
                i += 1
            exit(1)
        
        permisoPathAsistencia = bool(os.access(pathArchivoTxtAsistencia, os.W_OK))
        permisoPathDotacion = bool(os.access(pathArchivoTxtDotacion, os.W_OK))
        if not permisoPathAsistencia or not permisoPathDotacion:
            if not permisoPathAsistencia:
                pathNoPermisos.append(pathArchivoTxtAsistencia)
            if not permisoPathDotacion:
                pathNoPermisos.append(pathArchivoTxtDotacion)
            for permisos in pathNoPermisos:
                print('Error no tiene permisos de escritura en el directorio: {0}'.format(permisos))
            exit(1)

        archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoXlsAsistencia])
        if archivosValidos and encabezadosValidos:
            procesoAsistencia(fechaEntrada, archivoXlsAsistencia, pathArchivoTxtAsistencia)
        else:
            print("Error en Archivo: {0}".format(archivoXlsAsistencia))

        procesoDotacion(fechaEntrada, pathArchivoTxtDotacion)

    else:
        print("Error: El programa CRO necesita {0} parametros para su ejecucion".format(PROCESOS_GENERALES['DOTACION']['ARGUMENTOS_PROCESO']))


if __name__ == "__main__":
    main()
