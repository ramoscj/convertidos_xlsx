# Directorio de archivos de entrada XLSX
PATH_XLSX = 'INPUTS/'
# Directorio de archivos de salida TXT
PATH_TXT = 'OUTPUTS/'
# Directorio de archivos LOG de salida
PATH_LOG = 'PROCESO_LOG/'

FUGA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'FUGA',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 2,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_FUGA_AGENCIA',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'FUGA',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto) 
    'ENCABEZADO_XLSX': ['LPATTR_PER_RES', 'LLAVEA', 'LPATTR_COD_POLI', 'LPATTR_COD_ORIGEN', 'TIPO', 'LPATTR_COD_STAT', 'NRO_POLIZA', 'ESTADOPOLIZA',                 'FECHAINICIOVIGENCIA', 'RUT_CRO', 'NOMBRE_CRO', 'FECHAPROCESO', 'CONSIDERAR_FUGA'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'FUGA', 'STOCK_PROXIMO_MES', 'RUT', 'UNIDAD'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'LPATTR_PER_RES': 0, 'TIPO': 4, 'LPATTR_COD_STAT': 5, 'RUT_CRO': 9, 'CONSIDERAR_FUGA': 12}
}

ASISTENCIA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'ASISTENCIA',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 2,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Asistencia_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'ASISTENCIA',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto) 
    'ENCABEZADO_XLSX': ['EJECUTIVA', 'RUT', 'PLATAFORMA'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'VHC_MES', 'DIAS_HABILES_MES', 'VHC_APLICA', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'EJECUTIVA': 0, 'RUT': 1, 'PLATAFORMA': 2}
}

GESTION_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'GESTION',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': 'Gestión CRO',
    # Nombre del archivo PropietariosCRO XLSX que el proceso usara
    'ENTRADA_PROPIETARIOS_XLSX': 'Propietarios CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'GESTION',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto) 
    'ENCABEZADO_XLSX': ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO DE ÚLTIMA TAREA', 'ESTADO', 'FECHA DE LA ÚLTIMA MODIFICACIÓN', 'DUEÑO: NOMBRE COMPLETO'],
    # Nombre de las columnas del encabezado que tendra el archivo de PropietariosCRO (se usa para validar que el archivo este correcto) 
    'ENCABEZADO_PROPIETARIOS_XLSX': ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'FECHA DE LA ÚLTIMA MODIFICACIÓN', 'DUEÑO: NOMBRE COMPLETO', 'ASIGNADO A: NOMBRE COMPLETO', 'CUENTA: PROPIETARIO DEL CLIENTE: NOMBRE COMPLETO'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'ESTADO', 'ESTADO_UT', 'ID_CAMPANA', 'CAMPANA','RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'CAMPAÑA_ID': 0, 'FECHA_DE_CREACION': 1, 'NOMBRE_DE_CAMPAÑA': 2, 'ESTADO_UT': 3, 'ESTADO': 4},
    # Columnas que se utilizaran durante el procesamiento del archivo PropietariosCRO XLSX
    'COLUMNAS_PROPIETARIOS_XLSX': {'CAMPAÑA_ID': 0, 'DUEÑO_NOMBRE_COMPLETO': 4, 'ASIGNADO_NOMBRE_COMPLETO': 5, 'CUENTA_NOMBRE_COMPLETO': 6}
}

CAMPANHAS_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'CAMPANHA_ESPECIAL',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 2,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_CampañasEspeciales_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'PILOTO',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto) 
    'ENCABEZADO_XLSX': ['EJECUTIVA', 'RUT', 'PLATAFORMA', 'CANTIDAD GESTIONES CAMPAÑAS ESPECIALES'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'NUMERO_GESTIONES', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'RUT': 1, 'NUMERO_GESTIONES': 3 }
}

DOTACION_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'DOTACION',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 2,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Asistencia_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'ICOM_CA_CANAL_',
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['Rut', 'Nombres', 'Apellido Paterno', 'Apellido Materno', 'Direccion', 'Comuna', 'Telefono', 'Celular', 'Fecha Ingreso', 'Fecha Nacimiento', 'Fecha Desvinculacion', 'Correo Electronico', 'Rut Jefe ', 'Empresa', 'Sucursal', 'Cargo', 'Nivel Cargo', 'Canal Negocio', 'Rol Pago'],
}