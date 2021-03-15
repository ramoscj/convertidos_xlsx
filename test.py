from conexio_db import conectorDB
from diccionariosDB import periodoCampanasEjecutivos

# encabezadoBulk = ['id', 'id_periodo_ejecutivo', 'nombre_cliente', 'fecha_creacion', 'nombre_campana', 'numero_poliza', 'fecha_cierre', 'estado_retencion', 'estado_ut']
# if salidaInsertBulkCampanas(pathArchivoCsv, campanasPorPeriodo, encabezadoBulk):
#     if os.path.isfile(pathArchivoCsv):
#         pathActual = os.getcwd()
#         sql = """BULK INSERT proactiva_campanas_ejecutivos FROM '{}\{}' WITH (FORMAT = 'CSV', FIELDTERMINATOR = ',', ROWTERMINATOR = '\n', FIRSTROW = 2);"""
#         cursor.execute(sql.format(pathActual, pathArchivoCsv))
# db.commit()

# for valores in camapanasEjecutivos.values():
#     campanasPorPeriodo = []
#     idEjecutivo = valores['ID_EJECUTIVO']
#     if camapanasPeriodoEjecutivos.get(idEjecutivo):
#         campanasPorPeriodo = setearCampanasPorEjecutivo(valores['CAMPANAS'], camapanasPeriodoEjecutivos[idEjecutivo]['ID'])
#         sql = """INSERT INTO proactiva_campanas_ejecutivos (id_periodo_ejecutivo, nombre_cliente, fecha_creacion, nombre_campana, numero_poliza, fecha_cierre, estado_retencion, estado_ut) VALUES (?, ?, ?, ?, ?, ?, ?, ?);"""
#         print(len(campanasPorPeriodo))
#         cursor.executemany(sql, campanasPorPeriodo)
#         db.commit()

def actualizarPolizasReliquidadas(polizasReliquidadas, fechaProceso):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaActualizar = []
        for valores in polizasReliquidadas.values():
            polizasParaActualizar.append([fechaProceso, valores['POLIZA']])
        sql = """UPDATE retenciones_por_reliquidar SET fecha_reliquidacion = ? WHERE numero_poliza = ? AND fecha_reliquidacion IS NULL;"""
        cursor.executemany(sql, polizasParaActualizar)
        db.commit()
        return True
    except Exception as e:
        db.rollback()
        raise Exception('Error al insertar polizas para reliquidar: %s' % (e))
    finally:
        cursor.close()
        db.close()

x = dict()
x[1] = {'POLIZA': '344548'}
x[2] = {'POLIZA': '305338'}

# print(actualizarPolizasReliquidadas(x, '01/01/2021'))