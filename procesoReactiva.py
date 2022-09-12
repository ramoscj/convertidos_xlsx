import sys, os
import datetime

from config_xlsx import PATH_LOG, REACTIVA_CONFIG_XLSX, PROCESOS_GENERALES, PATH_RAIZ
from complementoCliente import COMPLEMENTO_CLIENTE_XLSX
from escribir_txt import salidaArchivoTxtProactiva, salidaLogTxt
from leerReactivaXLSX import LOG_PROCESO_REACTIVA, leerArchivoReactiva
from validaciones_texto import validaFechaInput, encontrarDirectorio, encontrarArchivo, compruebaEncabezado, sacarNombreArchivo

from crearXlsx import crearArchivoXlsx

def procesoReactiva(fechaInput, fechaRangoUno, fechaRangoDos, archivoXlsxInput, archivoCertificacionXls, archivoComplementoCliente, pathArchivoTxt, pathArchivoXlsx):
    
    hora = datetime.datetime.now()
    pathLogSalida = "REACTIVA/{0}log_{1}{2}_{3}.txt".format(PATH_LOG, 'REACTIVA', fechaInput, hora.strftime("%Y%m%d%H%M"))
    try:
        dataReactivaTxt, dataXlsx = leerArchivoReactiva(archivoXlsxInput, fechaInput, fechaRangoUno, fechaRangoDos, archivoCertificacionXls, archivoComplementoCliente)
        salidaXlsx = "{0}/{1}_{2}".format(pathArchivoTxt, fechaInput, REACTIVA_CONFIG_XLSX['SALIDA_XLSX'])
        logProceso = LOG_PROCESO_REACTIVA
        
        if dataReactivaTxt:
            for data in dataReactivaTxt:
                salidaTxt = "{0}/{1}{2}.txt".format(pathArchivoXlsx, data['NOMBRE_ARCHIVO'], fechaInput)
                if salidaArchivoTxtProactiva(salidaTxt, data['DATA'], data['ENCABEZADO']):
                    print("<a>&#128221;</a> Archivo TXT: {0} creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaTxt), len(data['DATA'])))
            archivoProduccionXslx = ['PRODUCCION_{0}'.format(fechaInput), REACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX_PERIODO'], dataXlsx]
            if crearArchivoXlsx(salidaXlsx, [archivoProduccionXslx]):
                print("<a>&#128221;</a> Archivo XLSX: {0}.xlsx creado con <strong> {1} registros</strong>".format(sacarNombreArchivo(salidaXlsx), len(dataXlsx)))
                
        else:
            print('<a style="color:red">Error no se crearon los Archivos para CORET REACTIVA</a>')

        if salidaLogTxt(pathLogSalida, logProceso):
            print("Archivo: {0} Creado!".format(pathLogSalida))

    except Exception as e:
        print(e)

def validarArchivosEntrada(archivosEntrada: []):
    
    archivosValidos = True
    encabezadosValidos = True
    encabezadosArchivos = [REACTIVA_CONFIG_XLSX['ENCABEZADO_XLSX'], COMPLEMENTO_CLIENTE_XLSX['ENCABEZADO'], REACTIVA_CONFIG_XLSX['ARCHIVO_BASE_CERTIFICACION']['ENCABEZADO']]
    coordenadasEncabezado = [REACTIVA_CONFIG_XLSX['COORDENADA_ENCABEZADO'], COMPLEMENTO_CLIENTE_XLSX['COORDENADA_ENCABEZADO'], REACTIVA_CONFIG_XLSX['ARCHIVO_BASE_CERTIFICACION']['COORDENADA_ENCABEZADO']]
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
    if len(sys.argv) == PROCESOS_GENERALES['REACTIVA']['ARGUMENTOS_PROCESO'] + 1:

        fechaEntrada = str(sys.argv[1])
        fechaRangoUno = str(sys.argv[2])
        fechaRangoDos = str(sys.argv[3])
        archivoReactivaXls = str(sys.argv[4])
        archivoCertificacionXls = str(sys.argv[5])
        archivoComplementoXls = str(sys.argv[6])
        pathArchivoTxt = str(sys.argv[7])
        pathArchivoXlsx = str(sys.argv[8])
        pathNoEncontrado = []
        directorioNumero = []
        pathNoPermisos = []

        if validaFechaInput(fechaEntrada):
            print("Fecha para el periodo %s OK!" % fechaEntrada)
        else:
            print("Fecha ingresada {0} incorrecta...".format(fechaEntrada))
            exit(1)

        salidaTxtDirectorio = encontrarDirectorio(pathArchivoTxt)
        salidaXlsxDirectorio = encontrarDirectorio(pathArchivoXlsx)
        if not salidaTxtDirectorio or not salidaXlsxDirectorio:
            if not salidaTxtDirectorio:
                pathNoEncontrado.append(pathArchivoTxt)
                directorioNumero.append(1)
            if not salidaXlsxDirectorio:
                pathNoEncontrado.append(pathArchivoXlsx)
                directorioNumero.append(2)
            i = 0
            for path in pathNoEncontrado:
                print('Error en el Directorio {0}: {1} no existe!'.format(directorioNumero[i], str(path)))
                i += 1
            exit(1)

        permisoPath = bool(os.access(pathArchivoTxt, os.W_OK))
        permisoPathXlsx = bool(os.access(pathArchivoXlsx, os.W_OK))
        if not permisoPath or not permisoPathXlsx:
            if not permisoPath:
                pathNoPermisos.append(pathArchivoTxt)
            if not permisoPathXlsx:
                pathNoPermisos.append(pathArchivoXlsx)
            for permisos in pathNoPermisos:
                print('Error no tiene permisos de escritura en el directorio: {0}'.format(permisos))
            exit(1)

        archivosValidos, encabezadosValidos = validarArchivosEntrada([archivoReactivaXls, archivoComplementoXls, archivoCertificacionXls])

        if archivosValidos and encabezadosValidos:
            procesoReactiva(fechaEntrada, fechaRangoUno, fechaRangoDos, archivoReactivaXls, archivoCertificacionXls, archivoComplementoXls, pathArchivoTxt, pathArchivoXlsx)
        else:
            print('<a style="color:red">ERROR EN ARCHIVOS DE ENTRADA.!</a>')
    else:
        print("Error: El programa REACTIVA necesita {0} parametros para su ejecucion".format(PROCESOS_GENERALES['REACTIVA']['ARGUMENTOS_PROCESO']))

if __name__ == "__main__":
    main()