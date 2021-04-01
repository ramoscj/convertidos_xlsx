import os.path
import sys

from config_xlsx import (ASISTENCIA_CONFIG_XLSX, CALIDAD_CONFIG_XLSX,
                         CAMPANHAS_CONFIG_XLSX, DOTACION_CONFIG_XLSX,
                         FUGA_CONFIG_XLSX, GESTION_CONFIG_XLSX, PATH_LOG,
                         PATH_TXT, PATH_XLSX)
from escribir_txt import salidaArchivoTxt, salidaLogTxt
from leerAsistenciaXLSX import LOG_PROCESO_ASISTENCIA, leerArchivoAsistencia
from leerCalidadXSLX import LOG_PROCESO_CALIDAD, leerArchivoCalidad
from leerCampanhasEspecialesXLSX import (LOG_PROCESO_CAMPANHAS,
                                         leerArchivoCampanhasEsp)
from leerDotacionXLSX import LOG_PROCESO_DOTACION, leerArchivoDotacion
from leerFugaXLSX import LOG_PROCESO_FUGA, leerArchivoFuga
from leerGestionXLSX import LOG_PROCESO_GESTION, leerArchivoGestion
from validaciones_texto import formatearFechaMesSiguiente, validaFechaInput


def procesoGeneral(procesoInput, fechaInput, archivoXlsxInput, pathArchivoTxt):

    if validaFechaInput(fechaInput):

        archivoXlsx = archivoXlsxInput
        pathLogSalida = ("%slog_%s%s.txt") % (PATH_LOG, procesoInput, fechaInput)

        if os.path.isfile(archivoXlsx):
            print("Archivo: %s encontrado." % archivoXlsx)
            print("Iniciando Lectura...")
            try:
                if procesoInput == "FUGA":
                    mesSiguiente = formatearFechaMesSiguiente(fechaInput)
                    salidaTxt = "%s%s.txt" % (
                        FUGA_CONFIG_XLSX['SALIDA_TXT'],
                        mesSiguiente.strftime("%Y%m"),
                    )
                    dataXlsx, encabezadoXlsx = leerArchivoFuga(archivoXlsx, fechaInput)
                    logProceso = LOG_PROCESO_FUGA

                elif procesoInput == "ASISTENCIA":
                    dataXlsx, encabezadoXlsx = leerArchivoAsistencia(
                        archivoXlsx, fechaInput
                    )
                    salidaTxt = "%s%s.txt" % (
                        ASISTENCIA_CONFIG_XLSX['SALIDA_TXT'],
                        fechaInput,
                    )
                    logProceso = LOG_PROCESO_ASISTENCIA

                elif procesoInput == "CAMPANHA_ESPECIAL":
                    dataXlsx, encabezadoXlsx = leerArchivoCampanhasEsp(
                        archivoXlsx, fechaInput
                    )
                    salidaTxt = "%s%s.txt" % (
                        CAMPANHAS_CONFIG_XLSX['SALIDA_TXT'],
                        fechaInput,
                    )
                    logProceso = LOG_PROCESO_CAMPANHAS

                elif procesoInput == "CALIDAD":
                    dataXlsx, encabezadoXlsx = leerArchivoCalidad(
                        archivoXlsx, fechaInput
                    )
                    salidaTxt = "%s%s.txt" % (
                        CALIDAD_CONFIG_XLSX['SALIDA_TXT'],
                        fechaInput,
                    )
                    logProceso = LOG_PROCESO_CALIDAD

                if dataXlsx:

                    archivoTxtOutput = '%s/%s' % (pathArchivoTxt, salidaTxt)
                    salidaArchivoTxt(archivoTxtOutput, dataXlsx, encabezadoXlsx)

                if salidaLogTxt(pathLogSalida, logProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)
            except Exception as e:
                print(e)
        else:
            print("Error: Archivo %s no encontrado" % archivoXlsx)


procesos = {
    "FUGA": FUGA_CONFIG_XLSX,
    "ASISTENCIA": ASISTENCIA_CONFIG_XLSX,
    "GESTION": GESTION_CONFIG_XLSX,
    "CAMPANHA_ESPECIAL": CAMPANHAS_CONFIG_XLSX,
    "CALIDAD": CALIDAD_CONFIG_XLSX,
    "DOTACION": DOTACION_CONFIG_XLSX
}
procesoInput = str(sys.argv[1]).upper()

def main():
    if procesos.get(procesoInput):
        if len(sys.argv) == procesos[procesoInput]["ARGUMENTOS_PROCESO"] + 1:
            if (
                procesoInput == "FUGA"
                or procesoInput == "ASISTENCIA"
                or procesoInput == "CAMPANHA_ESPECIAL"
                or procesoInput == "CALIDAD"
            ):
                fechaEntrada = str(sys.argv[2])
                archivoXls = str(sys.argv[3])
                pathArchivoTxt = str(sys.argv[4])
                procesoGeneral(procesoInput, fechaEntrada, archivoXls, pathArchivoTxt)

            elif procesoInput == "DOTACION":
                fechaEntrada = str(sys.argv[2])
                pathArchivoTxt = str(sys.argv[3])
                pathLogSalida = ("%slog_%s%s.txt") % (PATH_LOG, procesoInput, fechaEntrada)
                dataXlsxDotacion, encabezadoXlsxDotacion = leerArchivoDotacion(
                    fechaEntrada
                )
                salidaTxt = "%s%s.txt" % (
                    DOTACION_CONFIG_XLSX['SALIDA_TXT'],
                    fechaEntrada,
                )

                if dataXlsxDotacion:
                    archivoTxtOutput = '%s/%s' % (pathArchivoTxt, salidaTxt)
                    salidaArchivoTxt(archivoTxtOutput, dataXlsxDotacion, encabezadoXlsxDotacion)

                logProceso = LOG_PROCESO_DOTACION
                if salidaLogTxt(pathLogSalida, logProceso):
                    print("Archivo: %s creado !!" % pathLogSalida)

            elif procesoInput == "GESTION":
                fechaEntrada = str(sys.argv[2])
                fechaRangoUno = str(sys.argv[3])
                fechaRangoDos = str(sys.argv[4])
                archivoXls = str(sys.argv[5])
                archivoPropietariosXls = str(sys.argv[6])
                pathSalidaTxt = str(sys.argv[7])

                if os.path.isfile(archivoXls):
                    print("Archivo: %s encontrado." % archivoXls)
                    print("Iniciando Lectura...")
                    pathTxtSalida = PATH_TXT
                    archivoTxt = ("%s/%s%s.txt") % (
                        pathSalidaTxt,
                        GESTION_CONFIG_XLSX["SALIDA_TXT"],
                        fechaEntrada,
                    )
                    pathLogSalida = ("%slog_%s_%s.txt") % (
                        PATH_LOG,
                        procesoInput,
                        fechaEntrada,
                    )
                    dataXlsx, encabezadoXlsx = leerArchivoGestion(
                        archivoXls, fechaEntrada, fechaRangoUno, fechaRangoDos, archivoPropietariosXls
                    )
                    if dataXlsx and salidaArchivoTxt(archivoTxt, dataXlsx, encabezadoXlsx):
                        LOG_PROCESO_GESTION.setdefault(
                            "SALIDA_TXT",
                            {
                                len(LOG_PROCESO_GESTION)
                                + 1: "Archivo: %s creado!! " % archivoTxt
                            },
                        )
                    erroresProceso = LOG_PROCESO_GESTION
                    if salidaLogTxt(pathLogSalida, erroresProceso):
                        print("Archivo: %s creado !!" % pathLogSalida)
                else:
                    print("Error: Archivo %s no encontrado" % archivoXls)
        else:
            print(
                "Error: El programa "
                '"%s"'
                " necesita %s parametros para su ejecucion"
                % (procesoInput, procesos[procesoInput]["ARGUMENTOS_PROCESO"])
            )
    else:
        print('Error: Proceso "' "%s" '" no encontrado' % procesoInput)

if __name__ == "__main__":
    main()