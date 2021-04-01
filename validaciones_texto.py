import datetime
import string

from dateutil.relativedelta import relativedelta


def validaFechaInput(fecha_x):
    try:
        f1 = fecha_x[0:4]
        f2 = fecha_x[4:6]
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
            errorMsg = "Celda%s - validaFechaCelda: %s | %s" % (setearCelda(celdaFila), str(celdaFila.value), e)
            return errorMsg
    except Exception as e:
        errorMsg = "Celda%s - validaFechaCelda: %s | %s" % (setearCelda(celdaFila), str(celdaFila.value), e)
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
        errorMsg = "Celda%s - setearFechaCelda: %s | %s" % (setearCelda(celdaFila), str(celdaFila.value), e)
        return errorMsg

def setearFechaInput(fecha):
    try:
        fechaAnho = str(fecha)[0:4]
        fechaMes = str(fecha)[4:6]
        fechaDia = str(fecha)[6:8]
        fechaSalida = datetime.date(int(fechaAnho), int(fechaMes), int(fechaDia))
        return fechaSalida
    except Exception as e:
        errorMsg = "Error %s, setearFechaInput YYYYMMDD | %s" % (fecha, e)
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
                error = 'Celda%s - %s;Encabezado incorrecto;%s' % (celda, nombreArchivo, encabezadoXls[i])
                columnasError.setdefault(len(columnasError)+1, error)
            i += 1
    if len(columnasError) > 0:
        return columnasError
    else:
        return True

def setearCelda(celda):
    resto, separador, celdaN = str(celda).partition(".")
    return ('<%s') % celdaN

def setearCelda2(celda, cantidad, *nroRegistro):
    if cantidad > 0:
        if celda[cantidad].value is None:
            cantidad -= 1
            setearCelda2(celda, cantidad, nroRegistro[0])
        if cantidad == len(celda) -1:
            resto, separador, celdaN = str(celda[cantidad]).partition(".")
            coordenada = 'Celda<%s' % celdaN
        else:
            resto, separador, celdaN = str(celda[cantidad]).partition(".")
            letrasAbecedario = list(string.ascii_uppercase)
            coordenada = 'Celda<%s%s>' % (letrasAbecedario[len(celda)-1], nroRegistro[0])
    else:
        resto, separador, celdaN = str(celda).partition(".")
        coordenada = 'Celda<%s' % celdaN
    return coordenada

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
    if plataformaSalida != 'CRO':
        plataformaSalida = plataforma
    return plataformaSalida

def formatearFechaYM(fecha):
    try:
        fechaAnho = str(fecha)[0:4]
        fechaMes = str(fecha)[4:6]
        fechaSalida = datetime.date(int(fechaAnho), int(fechaMes), 1)
        return fechaSalida
    except Exception as e:
        errorMsg = "Error %s, formatearFechaYM YYYYMM | %s" % (fecha, e)
        raise Exception(errorMsg)

def formatearFechaMesSiguiente(fecha):
    try:
        fechaAnho = str(fecha)[0:4]
        fechaMes = str(fecha)[4:6]
        fechaSalida = datetime.datetime(int(fechaAnho), int(fechaMes), 1).replace(day=1).date()+relativedelta(months=1)
        return fechaSalida
    except Exception as e:
        errorMsg = "Error %s, formato correcto YYYYMM | %s" % (fecha, e)
        raise Exception(errorMsg)

def mesSiguienteUltimoDia(fecha):
    try:
        fechaAnho = str(fecha)[0:4]
        fechaMes = str(fecha)[4:6]
        fechaSalida = datetime.datetime(int(fechaAnho), int(fechaMes), 1).replace(day=1).date()+relativedelta(months=1)
        fechaSalida = datetime.datetime(int(fechaSalida.strftime("%Y")), int(fechaSalida.strftime("%m")), 1).replace(day=1).date()+relativedelta(months=1)+datetime.timedelta(days=-1)
        return fechaSalida
    except Exception as e:
        errorMsg = "Error %s, formato correcto YYYYMM | %s" % (fecha, e)
        raise Exception(errorMsg)

def separarNombreApellido(nombreCompleto):
    apellidos, separador, nombres = nombreCompleto.partition(',')
    nombresSalida = setearNombre(nombres.split())
    apellidoPaterno = apellidos.split()
    apellidoMaterno = setearApellidoMaterno(apellidos.split())
    return apellidoPaterno[0], apellidoMaterno, nombresSalida

def setearApellidoMaterno(texto):
    textoFormateado = ''
    for cantidad in range(1,len(texto)):
        if cantidad == len(texto)-1:
            textoFormateado += texto[cantidad]
        else:
            textoFormateado += texto[cantidad] + ' '
    return textoFormateado

def setearNombre(texto):
    textoFormateado = ''
    for cantidad in range(0,len(texto)):
        if cantidad == len(texto)-1:
            textoFormateado += texto[cantidad]
        else:
            textoFormateado += texto[cantidad] + ' '
    return textoFormateado

def formatearFechaMesAnterior(fecha):
    try:
        fechaAnho = str(fecha)[0:4]
        fechaMes = str(fecha)[4:6]
        fechaSalida = datetime.datetime(int(fechaAnho), int(fechaMes), 1).replace(day=1).date()-relativedelta(months=1)
        return fechaSalida
    except Exception as e:
        errorMsg = "Error %s, formato correcto YYYYMM | %s" % (fecha, e)
        raise Exception(errorMsg)

def formatearNumeroPoliza(numeroPoliza):
    polizaSalida = None
    if numeroPoliza is not None and numeroPoliza != '':
        nroPolizaSolo, separador, nroCertificado = str(numeroPoliza).partition("_")
        polizaSalida = nroPolizaSolo
    return polizaSalida

def formatearIdCliente(idCliente):
    campana, separador, codCliente = str(idCliente).partition(":")
    nombreCliente, separador, rutEjecutivo = str(codCliente).partition("-")
    return nombreCliente.strip()

def convertirALista(dataLista: dict):
    listaFinal = []
    for llave, valores in dataLista.items():
        lista = []
        for valor in valores:
            lista.append(dataLista[llave][valor])
        listaFinal.append(list(lista))
    return listaFinal

def convertirListaCampana(dataLista: dict, ejecutivosExistentes, fechaPeriodo):
    lista = []
    for valores in dataLista.values():
        if not ejecutivosExistentes.get(valores['ID_EJECUTIVO']):
            lista.append([valores['ID_EJECUTIVO'], fechaPeriodo])
    return lista

def setearCampanasPorEjecutivo(dataLista: [], idEjecutivo):
    lista = []
    for valores in dataLista:
        data = [idEjecutivo]
        data += valores
        lista.append(data)
    return lista

def convertirDataReact(dataReact: dict):
    dataFinal = dict()
    for valores in dataReact.values():
        pk = '{0}_{1}'.format(valores['ID_CAMPANA'], valores['POLIZA'])
        dataFinal[pk] = {'ESTADO_VALIDO_REACT': valores['ESTADO_VALIDO_REACT'], 'CONTACTO_REACT': valores['CONTACTO_REACT'], 'EXITO_REPETIDO_REACT': valores['EXITO_REPETIDO_REACT'], 'ID_EMPLEADO': valores['ID_EMPLEADO'], 'ID_CAMPANA': valores['ID_CAMPANA'], 'CAMPANA': valores['CAMPANA'], 'POLIZA': valores['POLIZA'], 'REPETICIONES': valores['REPETICIONES']}
    return dataFinal