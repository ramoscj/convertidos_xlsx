import datetime
from dateutil.relativedelta import relativedelta

def validaFechaInput(f1, f2, fecha_x):
    try:
        if type(datetime.date(int(f1), int(f2), 1)) is datetime.date and len(fecha_x) <= 6:
            return True
        else:
            return False
    except Exception as e:
        print("Error de fecha, formato correcto YYYYMMDD | %s" % e)

def validaFechaCelda(celdaFila):
    try:
        fecha = str(celdaFila.value)
        fechaAnho = fecha[0:4]
        fechaMes = fecha[4:6]
        if type(datetime.date(int(fechaAnho), int(fechaMes), 1)) is datetime.date:
            return celdaFila
        else:
            errorMsg = "Celda%s - Fecha incorrecta: %s | %s" % (setearCelda(celdaFila), str(celdaFila.value), e)
            return errorMsg
    except Exception as e:
        errorMsg = "Celda%s - Fecha incorrecta: %s | %s" % (setearCelda(celdaFila), str(celdaFila.value), e)
        return errorMsg

def setearFechaCelda(celdaFila):
    try:
        fecha = str(celdaFila.value).replace("-","")
        fechaAnho = fecha[0:4]
        fechaMes = fecha[4:6]
        fechaDia = fecha[6:8]
        fechaSalida = datetime.date(int(fechaAnho), int(fechaMes), int(fechaDia))
        return fechaSalida
    except Exception as e:
        errorMsg = "Celda%s - Fecha incorrecta: %s | %s" % (setearCelda(celdaFila), str(celdaFila.value), e)
        return errorMsg

def setearFechaInput(fecha):
    try:
        fechaAnho = str(fecha)[0:4]
        fechaMes = str(fecha)[4:6]
        fechaDia = str(fecha)[6:8]
        fechaSalida = datetime.date(int(fechaAnho), int(fechaMes), int(fechaDia))
        return fechaSalida
    except Exception as e:
        errorMsg = "Error %s, formato correcto YYYYMMDD | %s" % (fecha, e)
        # return errorMsg
        raise Exception(errorMsg)

def formatearRut(rut):
    rutMantisa, separador, dv = rut.partition("-")
    rutSalida = '%s%s' % (rutMantisa, dv)
    return rutSalida

def formatearRutGion(rut):
    caracteres = len(str(rut).strip())
    rutSalida = '%s-%s' % (rut[0:caracteres-1], rut[caracteres-1:caracteres])
    return rutSalida

def validarEncabezadoXlsx(filasXlsx: [], encabezadoXls: [], nombreArchivo):
    columnasError = dict()
    i = 0
    for fila in filasXlsx:
        for celda in fila:
            if str(celda.value).upper() != encabezadoXls[i]:
                celda = setearCelda(str(fila[i]))
                error = 'Celda%s - %s: Encabezado incorrecto: %s' % (celda, nombreArchivo, encabezadoXls[i])
                columnasError.setdefault(len(columnasError)+1, error)
            i += 1
    if len(columnasError) > 0:
        return columnasError
    else:
        return True

def setearCelda(celda):
    resto, separador, celdaN = str(celda).partition(".")
    return ('<%s') % celdaN

def primerDiaMes(fecha):
    fechaAnho = str(fecha)[0:4]
    fechaMes = str(fecha)[4:6]
    primerDia = datetime.datetime(int(fechaAnho), int(fechaMes), 1).replace(day=1).date()
    return primerDia

def ultimoDiaMes(fecha):
    fechaAnho = str(fecha)[0:4]
    fechaMes = str(fecha)[4:6]
    ultimoDia = datetime.datetime(int(fechaAnho), int(fechaMes), 1).replace(day=1).date()+relativedelta(months=1)+datetime.timedelta(days=-1)
    return ultimoDia

def formatearPlataformaCRO(plataforma):
    carateres = len(str(plataforma).strip())
    plataformaSalida = str(plataforma[0:carateres-1]).strip()
    return plataformaSalida

# print(setearFechaInput('20200101'))
# print(formatearPlataformaCRO('CRO 2'))
