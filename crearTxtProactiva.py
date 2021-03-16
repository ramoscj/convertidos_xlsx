import os.path
import sys

from config_xlsx import PATH_LOG, PATH_TXT, PATH_XLSX, PROACTIVA_CONFIG_XLSX
from escribir_txt import salidaArchivoTxtProactiva, salidaLogTxt
from leerProactivaXLSX import LOG_PROCESO_PROACTIVA, leerArchivoProactiva
from validaciones_texto import validaFechaInput

if __name__ == "__main__":
    if len(sys.argv) == 5:

        fechaProceso = str(sys.argv[1])
        archivoXls = str(sys.argv[2])
        archivoComplmentoCliente = str(sys.argv[3])
        PATH_TXT = str(sys.argv[4])

        if validaFechaInput(fechaProceso):

            # archivoXls = ("PROACTIVA/%s%s_%s.xlsx") % (
            #     PATH_XLSX,
            #     PROACTIVA_CONFIG_XLSX["ENTRADA_XLSX"],
            #     fechaProceso
            # )

            if os.path.isfile(archivoXls):

                dataReactivaTxt, encabezadoTxt = leerArchivoProactiva(
                    archivoXls, fechaProceso, archivoComplmentoCliente
                )
                archivoTxtSalida = "%s/%s%s.txt" % (
                    PATH_TXT,
                    PROACTIVA_CONFIG_XLSX["SALIDA_TXT"],
                    fechaProceso,
                )

                if dataReactivaTxt:
                    salidaArchivoTxtProactiva(archivoTxtSalida, dataReactivaTxt, encabezadoTxt)

                logProceso = LOG_PROCESO_PROACTIVA
                pathLogSalida = ("PROACTIVA/%slog_%s%s.txt") % (
                    PATH_LOG,
                    PROACTIVA_CONFIG_XLSX["SALIDA_TXT"],
                    fechaProceso,
                )
                if salidaLogTxt(pathLogSalida, logProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)
            else:
                print("Error: Archivo %s no encontrado" % archivoXls)
    else:
        print(
            "Error: El proceso "
            "'%s'"
            " necesita los parametros: "'FECHA_PERIODO'", "'ARCHIVO_PROACTIVA.XLSX'", "'ARCHIVO_COMPLEMENTO_CLIENTE.XLSX'", "'PATH_SALIDA_TXT'""
            % PROACTIVA_CONFIG_XLSX["PROCESO"]
        )
