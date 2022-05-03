from conexio_db import conectorDB
from diccionariosDB import listaEstadoUtDesc, estadoRetencionProDesc, listaEstadoUtContacto

def definirEstadoPro(estado):
    listaContactado = {1: 'Pendiente', 2: 'Terminado con Exito', 3: 'Terminado sin Exito'}
    salidaEstado = 'Sin Gestion'
    if listaContactado.get(estado):
        salidaEstado = listaContactado.get(estado)
    return salidaEstado

def validarClienteContacto(estado, estadoUt):
    listaEstadoContactado = listaEstadoUtContacto()
    contacto = 'NO CONTACTADO'
    if estado >= 1:
        if estado == 2:
            contacto = 'CONTACTADO'
        else:
            if listaEstadoContactado.get(estadoUt):
                contacto = 'CONTACTADO'
    
    return contacto

def definirEstadoUtPro(estadoUt):
    listaEstadoUt = listaEstadoUtDesc()
    salidaEstadoUt = None
    if listaEstadoUt.get(estadoUt):
        salidaEstadoUt = listaEstadoUt.get(estadoUt)
    return salidaEstadoUt
    
def definirBooleano(valor):
    salida = 'NO'
    if valor == 1:
        salida = 'SI'
    return salida

def definirEstadoRetencionPro(estadoRetencion):
    listaEstadoRetencion = estadoRetencionProDesc()
    salidaEstadoRet = None
    if listaEstadoRetencion.get(estadoRetencion):
        salidaEstadoRet = listaEstadoRetencion.get(estadoRetencion)
    return salidaEstadoRet

def dataXlsxReliquidacionesProactiva(periodo, dataReliquidaciones):
    try:
        db = conectorDB()
        cursor = db.cursor()
        polizasParaRequilidar = dict()
        sql = """SELECT fecha_creacion, nombre_campana_completo, id_ejecutivo, estado_pro, polizas_campana, fecha_cierre, numero_poliza, fecha_expiracion_coret, estado_retencion, campana_id, estado_ut_pro, fecha_ultimo_pago, fecha_mandato, estado_mandato, cobranza_rel_pro, pacpat_rel_pro FROM proactiva_campanas_ejecutivos LEFT JOIN proactiva_campanas_periodo_ejecutivos pcpe ON proactiva_campanas_ejecutivos.id_periodo_ejecutivo = pcpe.id WHERE pcpe.periodo = ? AND reliquidacion = 1 ORDER BY id_ejecutivo"""
        cursor.execute(sql, (periodo))
        for (fecha_creacion, nombre_campana_completo, id_ejecutivo, estado_pro, polizas_campana, fecha_cierre, numero_poliza, fecha_expiracion_coret, estado_retencion, campana_id, estado_ut_pro, fecha_ultimo_pago,fecha_mandato, estado_mandato, cobranza_rel_pro, pacpat_rel_pro) in cursor:
            
            pk = '{0}_{1}_{2}'.format(campana_id, id_ejecutivo, numero_poliza)
            if not dataReliquidaciones.get(pk):
                continue
            
            pagaCobranza = dataReliquidaciones[pk]['COBRANZA_REL_PRO']
            pagaMandato = dataReliquidaciones[pk]['PACPAT_REL_PRO']
            
            polizasParaRequilidar[pk] = {'FECHA_CREACION': fecha_creacion, 'NOMBRE_CAMPANA': nombre_campana_completo, 'CODIGO_EMPLEADO': id_ejecutivo, 'ESTADO_PRO': definirEstadoPro(estado_pro), 'POLIZAS_CAMPANA': polizas_campana, 'FECHA_CIERRE': fecha_cierre, 'POLIZA': str(numero_poliza), 'FECHA_EXPIRACION': fecha_expiracion_coret, 'ESTADO_RETENCION': definirEstadoRetencionPro(estado_retencion), 'CAMPAÑA_ID': campana_id, 'ESTADO_UT_PRO': definirEstadoUtPro(estado_ut_pro), 'FECHA_ULTPAGO': fecha_ultimo_pago, 'ESTADO_MANDATO': estado_mandato, 'FECHA_MANDATO': fecha_mandato, 'COBRANZA_PRO': definirBooleano(cobranza_rel_pro), 'PACPAT_PRO': definirBooleano(pacpat_rel_pro), 'PAGA_COBRANZA': definirBooleano(pagaCobranza), 'PAGA_MANDATO': definirBooleano(pagaMandato)}
        return polizasParaRequilidar
    except Exception as e:
        raise Exception('Error def dataXlsxProactiva(): %s' % e)
    finally:
        cursor.close()
        db.close()