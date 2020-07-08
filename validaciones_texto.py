import datetime

def validar_fecha(f1, f2, fecha_x):
    try:
        if type(datetime.date(int(f1), int(f2), 1)) is datetime.date and len(fecha_x) <= 6:
            return True
        else:
            return False
    except Exception as e:
        print("Error de fecha, formato correcto YYYYMM | %s" % e)