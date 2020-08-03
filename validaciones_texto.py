import datetime

def validaFechaInput(f1, f2, fecha_x):
    try:
        if type(datetime.date(int(f1), int(f2), 1)) is datetime.date and len(fecha_x) <= 6:
            return True
        else:
            return False
    except Exception as e:
        print("Error de fecha, formato correcto YYYYMMDD | %s" % e)

def setearFechaInput(fecha, coordenadaFila):
    try:
        buscarDv = fecha.find('-')
        if buscarDv > 0:
            fecha = fecha.replace("-","")
        fechaAnho = fecha[0:4]
        fechaMes = fecha[4:6]
        fechaDia = fecha[6:8]
        fechaSalida = datetime.date(int(fechaAnho), int(fechaMes), int(fechaDia))
        if type(fechaSalida) is datetime.date:
            return fechaSalida
        else:
            return False
    except Exception as e:
        print("Error de fecha, formato correcto DDMMYYYY | %s - %s" % (e, coordenadaFila))

def formatearRut(rut):
    rutMantisa, separador, dv = rut.partition("-")
    rutSalida = '%s%s' % (rutMantisa, dv)
    return rutSalida

def validarEncabezadoXlsx(filasXlsx: [], encabezadoXls: []):
    columnasError = []
    i = 0
    for fila in filasXlsx:
        for celda in fila:
            if str(celda.value).upper() != encabezadoXls[i]:
                celda = str(fila[i])
                resto, separador, celdaN = celda.partition(".")
                error = 'Error celda <%s: %s' % (celdaN, encabezadoXls[i])
                columnasError.append(error)
                print(error)
            i += 1
    if len(columnasError) > 0:
        return False
    else:
        return True

# print(setearFechaInput('20200101', 'x'))
