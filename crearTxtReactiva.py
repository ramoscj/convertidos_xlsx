import os.path
import sys

from config_xlsx import PATH_LOG, PATH_TXT, PATH_XLSX, REACTIVA_CONFIG_XLSX
from escribir_txt import salidaArchivoTxtProactiva, salidaLogTxt
from leerReactivaXLSX import LOG_PROCESO_REACTIVA, leerArchivoReactiva
from validaciones_texto import validaFechaInput

if __name__ == "__main__":
    if len(sys.argv) == 8:

        fechaProceso = str(sys.argv[1])
        fechaRangoUno = str(sys.argv[2])
        fechaRangoDos = str(sys.argv[3])
        archivoReactivaXls = str(sys.argv[4])
        archivoCertificacionXls = str(sys.argv[5])
        archivoComplementoXls = str(sys.argv[6])
        PATH_TXT = str(sys.argv[7])

        if validaFechaInput(fechaProceso):

            if os.path.isfile(archivoReactivaXls) and os.path.isfile(archivoCertificacionXls) and os.path.isfile(archivoComplementoXls):

                print("Archivos: {0}, {1}, {2} encontrados.".format(archivoReactivaXls, archivoCertificacionXls, archivoComplementoXls))
                print("Iniciando Lectura...")
                dataReactivaTxt = leerArchivoReactiva(
                    archivoReactivaXls, fechaProceso, fechaRangoUno, fechaRangoDos, archivoCertificacionXls, archivoComplementoXls
                )

                if dataReactivaTxt:
                    for data in dataReactivaTxt:
                        archivoTxtSalida = "{0}/{1}{2}.txt".format(
                            PATH_TXT,
                            data['NOMBRE_ARCHIVO'],
                            fechaProceso
                        )
                        salidaArchivoTxtProactiva(archivoTxtSalida, data['DATA'], data['ENCABEZADO'])

                logProceso = LOG_PROCESO_REACTIVA
                pathLogSalida = ("{0}\log_REACTIVA{1}.txt").format(
                    REACTIVA_CONFIG_XLSX["PATH_LOG"],
                    fechaProceso
                )
                if salidaLogTxt(pathLogSalida, logProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)
            else:
                print("Archivos: {0}, {1}, {2} NO encontrados.".format(archivoReactivaXls, archivoCertificacionXls, archivoComplementoXls))
    else:
        print(
            "Error: El proceso "
            "'%s'"
            " necesita los parametros: "'FECHA_PERIODO'", "'FECHA_RANGOUNO'", "'FECHA_RANGODOS'", "'ARCHIVO_REACTIVA.XLSX'", "'ARCHIVO_BASE_CERTIFICACION.XLSX'", "'ARCHIVO_COMPLEMENTO_CLIENTE.XLSX'", "'PATH_SALIDA_TXT'""
            % REACTIVA_CONFIG_XLSX["PROCESO"]
        )
