# from imp import load_source
from pandas import read_excel
import pandas as pd
import numpy as np
from alive_progress import alive_bar
from diccionariosDB import buscarCamphnasDb, buscarRutEjecutivosDb, listaEstadoUtCro, listaEstadoUtContactoCro
from validaciones_texto import (primerDiaMes, setearCelda, setearFechaCelda,
                                setearFechaInput, ultimoDiaMes, primerDiaMes,
                                setearCelda2)
from escribir_txt import escribirArchivoTxt
from leerGestionXLSX import extraerPropietariosCro


file_name = r'CRO\INPUTS\202112_Gestion_CRO.xlsx'
prpietarios = r'CRO\INPUTS\202112_Propietarios_Inbound_CRO.xlsx'
# dataPropietarios = {'a5e3w000001Tx2o': {'ID_EMPLEADO' : 1233}}
# dataPropietarios = extraerPropietariosCro(prpietarios)
listaEstadoUt = listaEstadoUtCro()

def definirFecha(fila):
    fecha = fila['FECHA']
    
    pass

def limpiarDataPropietarios(archivo):
    df = read_excel(archivo)
    df.columns = ['CAMPANA_ID', 'FECHA_CREACION', 'NOMBRE_CAMPANA', 'ESTADO_UT', 'ESTADO', 'FECHA', 'ID_EMPLEADO']
    df = df.dropna(subset=['ID_EMPLEADO'])
    df2 = df.copy()
    # dfFechaFinal = df2.index[ df2.groupby('CAMPANA_ID').agg({ 'FECHA': 'max' }).agg ]
    
    # df2['DUPLICADO'] = df2.duplicated(subset=['CAMPANA_ID'], keep='first')
    # df2 = df[ (df2['DUPLICADO'] == False) ]
    # resultado = pd.merge(df2[['CAMPANA_ID', 'FECHA_CREACION', 'NOMBRE_CAMPANA', 'ESTADO_UT', 'ESTADO']], dfFechaFinal[['FECHA', 'ID_EMPLEADO']], how='left', on='CAMPANA_ID')
    # df2['FECHA'] = dfFechaFinal['FECHA']
    # campanasDuplicadas = df[ (df2['DUPLICADO'] == )]
    # df2.drop()
    # df2 = df.apply(definirFecha, axis=1)
    # return df[ (df2['DUPLICADO'] == True) ]
    
    # d = df2.sort_values('FECHA', ascending=False)
    # d = d.drop_duplicates(['CAMPANA_ID'])
    
    fechasMaximas = df2.groupby(['CAMPANA_ID']).FECHA.transform(max)
    df2 = df2.loc[df2.FECHA == fechasMaximas]
    df2 = df2.drop_duplicates(['CAMPANA_ID'])
    # print(df2)
    return df2

def getEstado(fila):
    listaEstado = {'Pendiente': 1, 'Terminado con Exito': 2, 'Terminado sin Exito': 3}
    if listaEstado.get(fila['ESTADO']):
        estadoFinal = listaEstado[fila['ESTADO']]
    elif fila['ESTADO'] == 'Sin Gestion':
        estadoFinal = 0
    else:
        print(fila['CAMPANA_ID'])
    return estadoFinal

def getEstadoUt(fila):
    estadoUt = str(fila['ESTADO_UT']).upper()
    if listaEstadoUt.get(estadoUt):
        return listaEstadoUt[estadoUt]
    else:
        return 0

def cambioDeEmpleado2(fila, dataPropietarios):
    ejecutivosExistentesDb = buscarRutEjecutivosDb(fila['FIN_PERIODO'], fila['INICIO_PERIODO'])
    ejecutivoCorrecto = fila['ID_EMPLEADO']
    if dataPropietarios.get(fila['CAMPANA_ID']):
        ejecutivoCorrecto = str(dataPropietarios[fila['CAMPANA_ID']]['ID_EMPLEADO'])
    
    if not ejecutivosExistentesDb.get(ejecutivoCorrecto):
        ejecutivoCorrecto = None
        
    return ejecutivoCorrecto

def cambioDeEmpleado(fila, ejecutivosExistentesDb):
    
    ejecutivoCorrecto = fila['EJECUTIVO_CORRECTO']
    if not ejecutivosExistentesDb.get(ejecutivoCorrecto):
        ejecutivoCorrecto = None
        
    return ejecutivoCorrecto

def empleadoPropietario(fila, dataPropietarios):
    ejecutivoPropietario = None
    if dataPropietarios.get(fila['CAMPANA_ID']):
        ejecutivoPropietario = str(dataPropietarios[fila['CAMPANA_ID']]['ID_EMPLEADO'])
        
    return ejecutivoPropietario

def definirContacto(fila):
    listaEstadoContactado = listaEstadoUtContactoCro()
    estadoUt = str(fila['ESTADO_UT']).upper()
    contacto = 'NO'
    if listaEstadoContactado.get(estadoUt) or fila['ESTADO_FINAL'] == 2:
        contacto = 'SI'
    return contacto
    

# if __name__ == "__main__":
#     fechaInicioPeriodo = setearFechaInput('20211124')
#     fechaFinPeriodo = setearFechaInput('20211231')
#     fechaIncioMes = primerDiaMes('202112')
#     fechaFinMes = ultimoDiaMes('202112')
#     ejecutivosExistentesDb = buscarRutEjecutivosDb(fechaFinMes, fechaIncioMes)
#     with alive_bar(2) as bar:
        
#         # dataPropietarios = extraerPropietariosCro(prpietarios)
#         dataPropietarios = limpiarDataPropietarios(prpietarios)
#         bar()
        
#         columnaSalida = ['CRR', 'ESTADO_FINAL', 'ESTADO_UT_FINAL', 'CAMPANA_ID', 'NOMBRE_CAMPANA', 'EJECUTIVO_CORRECTO']
#         df = read_excel(file_name)
#         # Se eliminan estados NULL
#         df.columns = ['CAMPANA_ID', 'FECHA_CREACION', 'NOMBRE_CAMPANA', 'ESTADO_UT', 'ESTADO', 'FECHA_CIERRE', 'ID_EMPLEADO']
#         # df['INICIO_PERIODO'] = fechaIncioMes
#         # df['FIN_PERIODO'] = fechaFinMes
#         df = df.dropna(subset=['ESTADO'])
#         # Se eliminan fechas de cierre NULL
#         df = df.dropna(subset=['FECHA_CIERRE'])
#         # Se eliminan estados que no sean validos
#         estadosNoValidos = df[ (df['ESTADO'] != 'Pendiente') & (df['ESTADO'] != 'Terminado con Exito') & (df['ESTADO'] != 'Terminado sin Exito') & (df['ESTADO'] != 'Sin Gestion') ].index
#         df.drop(estadosNoValidos , inplace=True)
        
#         dataFiltrada = df[ (df['NOMBRE_CAMPANA'] == 'Inbound CRO')]
#         dataFiltradaPorMes = dataFiltrada[ (pd.to_datetime(dataFiltrada['FECHA_CIERRE']).dt.date >= fechaIncioMes) & (pd.to_datetime(dataFiltrada['FECHA_CIERRE']).dt.date <= fechaFinMes) ]
        
#         df2 = dataFiltradaPorMes.copy()
#         # df2['ID_EMPLEADO'] = df2['ID_EMPLEADO'].fillna(0)
#         # df2['ID_EMPLEADO_PROPIETARIO'] = df2.apply(empleadoPropietario, dataPropietarios = dataPropietarios, axis = 1)
#         merged = df2.merge(dataPropietarios[['CAMPANA_ID','ID_EMPLEADO']], on='CAMPANA_ID', how='left', suffixes=('_DATA', '_PROPIETARIO'))
#         merged['EJECUTIVO_CORRECTO'] = merged['ID_EMPLEADO_PROPIETARIO'].where(merged['ID_EMPLEADO_PROPIETARIO'].notnull(), merged['ID_EMPLEADO_DATA'])
        
#         merged['EJECUTIVO_CORRECTO'] = merged['EJECUTIVO_CORRECTO'].astype(int)
#         merged['EJECUTIVO_CORRECTO'] = merged['EJECUTIVO_CORRECTO'].astype(str)
#         merged['EJECUTIVO_CORRECTO'] = merged.apply(cambioDeEmpleado, ejecutivosExistentesDb = ejecutivosExistentesDb, axis = 1)
#         # merged = merged.dropna(subset=['ESTADO'])
        
#         merged['ESTADO_FINAL'] = merged.apply(getEstado, axis = 1)
#         merged['ESTADO_UT_FINAL'] = merged.apply(getEstadoUt, axis = 1)
#         merged['CONTACTO_FINAL'] = merged.apply(definirContacto, axis = 1)
        
        
#         merged = merged.dropna(subset=['EJECUTIVO_CORRECTO'])
#         merged['CRR'] = np.arange(1, len(merged)+1)
#         merged = merged[columnaSalida]
        
#         escribirArchivoTxt(r'CRO\OUTPUTS\test.txt',merged.values.tolist(), columnaSalida)
        
#         # dataPropietarios = limpiarDataPropietarios(prpietarios)
#         # dataPropietarios.to_csv('text.csv')
#         bar()

dataPropietarios = extraerPropietariosCro(prpietarios)