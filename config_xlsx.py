# Directorio de archivos de entrada XLSX
PATH_XLSX = 'INPUTS/'
# Directorio de archivos de salida TXT
PATH_TXT = 'OUTPUTS/'
# Directorio de archivos LOG de salida
PATH_LOG = 'PROCESO_LOG/'

# Parametros de conexion a la DB
ACCESO_DB = {
    'SERVIDOR': 'DESKTOP-8R9ENHE',
    'NOMBRE_DB': 'icom',
    'USUARIO': 'sa',
    'CLAVE': '5325106'
}

FUGA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'FUGA',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A1:P1',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_FUGA_AGENCIA',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'FUGA',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['LPATTR_PER_RES', 'LLAVEA', 'LPATTR_COD_POLI', 'LPATTR_COD_ORIGEN', 'TIPO', 'LPATTR_COD_STAT', 'TIPO_FUGA', 'NRO_POLIZA', 'ESTADOPOLIZA', 'FECHAINICIOVIGENCIA', 'ID_EMPLEADO', 'CODIGOPLAN', 'TIPOPRODUCTO', 'PRODUCTO', 'FECHAPROCESO', 'CONSIDERAR'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_FUGA_TXT': ['CRR', 'FUGA_MES_ANTERIOR', 'STOCK', 'RUT', 'UNIDAD'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'LPATTR_PER_RES': 0, 'TIPO': 4, 'LPATTR_COD_STAT': 5, 'ID_EMPLEADO': 10, 'CONSIDERAR_FUGA': 15}
}

ASISTENCIA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'ASISTENCIA',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A2:B2',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Asistencia_Plataformas',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'ASISTENCIA',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['ID EMPLEADO', 'PLATAFORMA'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'VHC_MES', 'DIAS_HABILES_MES', 'VHC_APLICA', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'ID_EMPLEADO': 0, 'PLATAFORMA': 1}
}

GESTION_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'GESTION',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A1:G1',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 7,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': 'Gestión CRO',
    # Nombre del archivo PropietariosCRO XLSX que el proceso usara
    'ENTRADA_PROPIETARIOS_XLSX': 'Propietarios CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'GESTION',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO DE ÚLTIMA TAREA', 'ESTADO', 'FECHA DE CIERRE', 'CAMPAIGN MEMBER: ID EMPLEADO'],
    # Nombre de las columnas del encabezado que tendra el archivo de PropietariosCRO (se usa para validar que el archivo este correcto)
    'ENCABEZADO_PROPIETARIOS_XLSX': ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO DE ÚLTIMA TAREA', 'ESTADO', 'FECHA', 'ACTIVITIES: EMPLOYEE ID'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'ESTADO', 'ESTADO_UT', 'ID_CAMPANA', 'CAMPANA','RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'CAMPAÑA_ID': 0, 'FECHA_DE_CREACION': 1, 'NOMBRE_DE_CAMPAÑA': 2, 'ESTADO_UT': 3, 'ESTADO': 4, 'FECHA_DE_CIERRE': 5, 'ID_EMPLEADO': 6},
    # Columnas que se utilizaran durante el procesamiento del archivo PropietariosCRO XLSX
    'COLUMNAS_PROPIETARIOS_XLSX': {'CAMPAÑA_ID': 0, 'FECHA': 5, 'ID_EMPLEADO': 6}
}

CAMPANHAS_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'CAMPANHA_ESPECIAL',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A2:C2',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_CampañasEspeciales_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'PILOTO',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['ID EMPLEADO', 'PLATAFORMA', 'CANTIDAD GESTIONES CAMPAÑAS ESPECIALES'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'NUMERO_GESTIONES', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'ID_EMPLEADO': 0, 'NUMERO_GESTIONES': 2 }
}

CALIDAD_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'CALIDAD',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A2:C2',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Calidad_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'CALIDAD',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['ID EMPLEADO', 'PLATAFORMA', 'CALIDAD'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'CALIDAD', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'ID_EMPLEADO': 0, 'CALIDAD': 2 }
}

DOTACION_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'DOTACION',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 3,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Asistencia_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'ICOM_CA_MTLFCC_',
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['Rut', 'Nombres', 'Apellido Paterno', 'Apellido Materno', 'Direccion', 'Comuna', 'Telefono', 'Celular', 'Fecha Ingreso', 'Fecha Nacimiento', 'Fecha Desvinculacion', 'Correo Electronico', 'Rut Jefe ', 'Empresa', 'Sucursal', 'Cargo', 'Nivel Cargo', 'Canal Negocio', 'Rol Pago'],
}

PROACTIVA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'PROACTIVA',
    # Coordenada del encabezado
    'COORDENADA_ENCABEZADO': 'A1:K1',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': 'Gestión CoRet Proactiva',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'ICOM_GESTION_CORPRO',
    # Nombre del archivo para POLIZAS RELIQUIDADAS
    'SALIDA_RELIQUIDACION': 'ICOM_RELIQUIDACION_CORPRO',

    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'COBRANZA_PRO', 'PACPAT_PRO', 'ESTADO_PRO', 'ESTADO_UT_PRO', 'REPETICION_PRO', 'ESTADO_RETENCION_PRO', 'RUT', 'ID_CAMPANA', 'CAMPANA', 'POLIZA'],

    # Nombre de las columnas de encabezado que tendra el archivo de POLIZAS RELOQUIDADAS
    'ENCABEZADO_RELIQUIDACIONES': ['CRR', 'COBRANZA_REL_PRO', 'PACPAT_REL_PRO', 'RUT', 'ID_CAMPANA', 'CAMPANA', 'POLIZA'],

    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'CAMPAIGN MEMBER: ID EMPLEADO', 'ESTADO', 'PÓLIZAS EN CAMPAÑA', 'FECHA DE CIERRE', 'PÓLIZA: NUMERO DE PÓLIZA', 'FECHA DE EXPIRACIÓN DEL CO-RET', 'ESTADO DE RETENCIÓN', 'MIEMBRO DE CAMPAÑA ID.', 'ESTADO DE ÚLTIMA TAREA'],

    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'FECHA_CREACION': 0, 'NOMBRE_DE_CAMPAÑA': 1, 'ID_EMPLEADO': 2, 'ESTADO': 3, 'FECHA_CIERRE': 5, 'NRO_POLIZA': 6, 'EXPIRACION_CORET': 7, 'ESTADO_RETENCION': 8, 'CAMAPAÑA_ID': 9, 'ESTADO_ULTIMA_TAREA': 10},

    # Estado de mandatos validos para aprobacion retenciones por ACTIVACION
    'ESTADO_MANDATO_VALIDO': {'APROBADO ENTIDAD RECAUDADORA': 1, 'APROBADO POR RECAUDACION': 1, 'APROBADO POR ENTIDAD RECAUDADORA': 1}
}

REACTIVA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'REACTIVA',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A1:L1',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 8,

    # Nombre del archivo base de certificacion
    'ARCHIVO_BASE_CERTIFICACION': {

        # Archivo de entrada
        'NOMBRE_ARCHIVO': 'Base Certificacion',

        # Nombre de las columnas del encabezado que tendra el archivo de Base Certificacion
        'ENCABEZADO': ['PÓLIZA', 'REQUERIMIENTO', 'FECHA DE LLAMADO', 'HORA DE LLAMADO', 'ID EMPLEADO', 'CANAL', 'SE ENCONTRÓ GRABACIÓN EN NICE', 'CLIENTE CONTACTADO', 'UTILIZA ARGUMENTO DE DESACTIVACIÓN DE MEDIO DE PAGO', 'CONFIRMA NOMBRE COMPLETO DE CLIENTE ', 'CONFIRMA RUT DEL CLIENTE', 'CONFIRMA TELÉFONO DE CONTACTO', 'EJECUTIVA DEJA SIN EFECTO CARTA EN FORMA VOLUNTARIA CLIENTE', 'MENCIONA NÚMERO DE PÓLIZA Y NOMBRE DEL PRODUCTO', 'MENCIONA QUE MANTENDRÁ COBERTURA VIGENTE', 'ENVÍA CARTA DE REVOCACIÓN ', 'EXPLICA COMO COMPLETAR CARTA DE REVOCACIÓN', 'TIPO DE CERTIFICACIÓN', 'FECHA CERTIFICACIÓN', 'EMAIL ENVIADO', 'GESTIÓN', 'UF', 'CARTA', 'EXISTE MIEMBRO DE CAMPAÑA - TERMINADO CON ÉXITO EN SALESFORCE'],

        # # Columnas que se utilizaran durante el procesamiento del archivo Base de Certificacion
        'COLUMNAS': {'NRO_POLIZA': 0, 'FECHA_LLAMADO': 2, 'ID_EMPLEADO': 4, 'CANAL': 5, 'TIPO_CERTIFICACION': 17}
    },

    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO', 'PÓLIZAS EN CAMPAÑA', 'FECHA DE CIERRE', 'PÓLIZA: NUMERO DE PÓLIZA', 'ESTADO DE RETENCIÓN', 'ÚLTIMA FECHA DE ACTIVIDAD', 'MIEMBRO DE CAMPAÑA ID.', 'ESTADO DE ÚLTIMA TAREA', 'ES LLAMADA SALIENTE', 'CAMPAIGN MEMBER: ID EMPLEADO'],

    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'FECHA_CREACION': 0, 'ESTADO': 2, 'FECHA_CIERRE': 4, 'NRO_POLIZA': 5, 'ESTADO_RETENCION': 6, 'CAMAPAÑA_ID': 8, 'ESTADO_ULTIMA_TAREA': 9, 'LLAMADA_SALIENTE': 10, 'ID_EMPLEADO': 11},

    # Estados de Retencion
    'ESTADOS_RETENCION': {'Cliente no vigente': 1, 'Cliente al día': 2, 'Lo va a pensar': 3, 'Solicita contacto por correo': 4, 'Pendiente de endoso': 5, 'Mantiene su producto': 6, 'Desiste el producto': 7, 'Anulado por cambio de producto': 8, 'Sin gestión por cierre': 9, 'Término programado de producto': 10, 'Queda vigente sin pagar': 11, 'Espera de carta de anulación': 12},

    # Configuracion de archivos de salida
    'SALIDA_TXT':{
        'GESTION':{
            'NOMBRE_SALIDA': 'GESTION_REACT',
            'ENCABEZADO': ['CRR', 'ESTADO_VALIDO_REACT', 'CONTACTO_REACT', 'EXITO_REPETIDO_REACT', 'REPETICION_REACT', 'RUT', 'ID_CAMPANA', 'CAMPANA', 'POLIZA']
        },
        'POLIZA':{
            'NOMBRE_SALIDA': 'POLIZA_REACT',
            'ENCABEZADO': ['CRR', 'ESTADO_POLIZA_REACT','POLIZA']
        },
        'CERTIFICACION':{
            'NOMBRE_SALIDA': 'CERTIFICACION_REACT',
            'ENCABEZADO': ['CRR', 'GRAB_CERTIFICADA_REACT', 'RUT', 'CAMPANA', 'POLIZA']
        },
    },

    # Directorio del LOG
    'PATH_LOG': 'REACTIVA\PROCESO_LOG',
}

COMPLEMENTO_CLIENTE_XLSX = {
    # Nombre del archivo Complemento Cliente
    'NOMBRE_ARCHIVO': 'COMPLEMENTO CLIENT vLite',

    # Nombre de las columnas del encabezado que tendra el archivo de Complemento Cliente
    'ENCABEZADO': ['NROPOLIZA', 'NROCERT', 'ESTADOPOLIZA', 'FEC_ULT_PAG', 'ESTADO_MANDATO', 'FECHA_MANDATO', 'FECHAPROCESO'],

    # # Columnas que se utilizaran durante el procesamiento del archivo Complemento Cliente
    'COLUMNAS': {'NRO_POLIZA': 0, 'NRO_CERT': 1, 'ESTADO_POLIZA': 2, 'FEC_ULT_PAG': 3, 'ESTADO_MANDATO': 4, 'FECHA_MANDATO': 5},
}