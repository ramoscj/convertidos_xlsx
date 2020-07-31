import datetime

def validaFechaInput(f1, f2, fecha_x):
    try:
        if type(datetime.date(int(f1), int(f2), 1)) is datetime.date and len(fecha_x) <= 6:
            return True
        else:
            return False
    except Exception as e:
        print("Error de fecha, formato correcto YYYYMMDD | %s" % e)

def setearFechaInput(fecha):
    try:
        fechaAnho = fecha[0:4]
        fechaMes = fecha[4:6]
        fechaDia = fecha[6:8]
        fechaSalida = datetime.date(int(fechaAnho), int(fechaMes), int(fechaDia))
        if type(fechaSalida) is datetime.date:
            return fechaSalida.strftime("%d/%m/%Y")
        else:
            return False
    except Exception as e:
        print("Error de fecha, formato correcto YYYYMM | %s" % e)

setearFechaInput('20201220')