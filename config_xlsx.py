# Directorio de archivos de entrada XLSX
PATH_XLSX = 'INPUTS/'
# Directorio de archivos de salida TXT
PATH_TXT = 'OUTPUTS/'
# Directorio de archivos LOG de salida
PATH_LOG = 'PROCESO_LOG/'

# Parametros de conexion a la DB
ACCESO_DB = {
    'SERVIDOR': 'DESKTOP-8R9ENHE',
    # 'SERVIDOR': 'database-2.caftetfeht1y.us-west-2.rds.amazonaws.com',
    'NOMBRE_DB': 'icom',
    'USUARIO': 'sa',
    'CLAVE': '5325106'
}

# Lista de ESTADO_ULTIMA_TAREA Contactado
listaEstadoContactado = {'Cliente retenido': 1, 'Llamado reprogramado': 2, 'Cliente no retenido': 3, 'No quiere escuchar': 4, 'Contacto con el asesor': 5, 'Apoyo del asesor al ejecutivo': 6, 'Pendiente respuesta cliente': 7, 'Carta de revocación pendiente': 8, 'Contacto por correo': 9, 'Campaña exitosa': 10, 'Solicita renuncia': 11, 'Cliente desconoce venta': 12, 'No se pudo instalar mandato': 13, 'Anulado por cambio de producto Metlife': 14, 'Cliente vive en el extranjero': 15, 'Cliente activa mandato': 16, 'Queda vigente sin pagar': 17, 'Lo está viendo con Asesor': 18}

listaEstadoNoContactado = {'Numero invalido': 1, 'Sin respuesta': 2, 'Buzón de voz': 3, 'Pagos al día': 4, 'Teléfono ocupado': 5, 'Teléfono apagado': 6, 'Campaña completada con 5 intentos': 7, 'Número equivocado': 8, 'Sin gestión de cierre': 9, 'Sin teléfono registrado': 10, 'Cliente no actualizado': 11, 'Temporalmente fuera de servicio': 12, 'Plazo previsto del producto': 13}

FUGA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'FUGA',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_FUGA_AGENCIA',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'FUGA',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['LPATTR_PER_RES', 'LLAVEA', 'LPATTR_COD_POLI', 'LPATTR_COD_ORIGEN', 'TIPO', 'LPATTR_COD_STAT', 'NRO_POLIZA', 'ESTADOPOLIZA', 'FECHAINICIOVIGENCIA', 'RUT_CRO', 'NOMBRE_CRO', 'FECHAPROCESO', 'CONSIDERAR'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_FUGA_TXT': ['CRR', 'FUGA_MES_ANTERIOR', 'STOCK', 'RUT', 'UNIDAD'],
    # 'ENCABEZADO_STOCK_TXT': ['CRR','STOCK_PROXIMO_MES', 'RUT', 'UNIDAD'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'LPATTR_PER_RES': 0, 'TIPO': 4, 'LPATTR_COD_STAT': 5, 'RUT_CRO': 9, 'CONSIDERAR_FUGA': 12}
}

ASISTENCIA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'ASISTENCIA',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Asistencia_Plataformas',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'ASISTENCIA',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['EJECUTIVO', 'NOMBRE POR SISTEMA', 'RUT', 'PLATAFORMA'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'VHC_MES', 'DIAS_HABILES_MES', 'VHC_APLICA', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'EJECUTIVO': 0, 'NOMBRE_RRH': 1, 'RUT': 2, 'PLATAFORMA': 3}
}

GESTION_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'GESTION',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 7,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': 'Gestión CRO',
    # Nombre del archivo PropietariosCRO XLSX que el proceso usara
    'ENTRADA_PROPIETARIOS_XLSX': 'Propietarios CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'GESTION',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO DE ÚLTIMA TAREA', 'ESTADO', 'FECHA DE LA ÚLTIMA MODIFICACIÓN', 'DUEÑO: NOMBRE COMPLETO'],
    # Nombre de las columnas del encabezado que tendra el archivo de PropietariosCRO (se usa para validar que el archivo este correcto)
    'ENCABEZADO_PROPIETARIOS_XLSX': ['MIEMBRO DE CAMPAÑA ID.', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'FECHA', 'DUEÑO: NOMBRE COMPLETO', 'ASIGNADO A: NOMBRE COMPLETO', 'CUENTA: PROPIETARIO DEL CLIENTE: NOMBRE COMPLETO'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'ESTADO', 'ESTADO_UT', 'ID_CAMPANA', 'CAMPANA','RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'CAMPAÑA_ID': 0, 'FECHA_DE_CREACION': 1, 'NOMBRE_DE_CAMPAÑA': 2, 'ESTADO_UT': 3, 'ESTADO': 4, 'FECHA_ULTIMA_MODF': 5, 'NOMBRE_COMPLETO': 6},
    # Columnas que se utilizaran durante el procesamiento del archivo PropietariosCRO XLSX
    'COLUMNAS_PROPIETARIOS_XLSX': {'CAMPAÑA_ID': 0, 'FECHA': 3, 'DUEÑO_NOMBRE_COMPLETO': 4, 'ASIGNADO_NOMBRE_COMPLETO': 5, 'CUENTA_NOMBRE_COMPLETO': 6}
}

CAMPANHAS_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'CAMPANHA_ESPECIAL',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
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

CALIDAD_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'CALIDAD',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '_Calidad_CRO',
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'CALIDAD',
    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['EJECUTIVA', 'RUT', 'PLATAFORMA', 'CALIDAD'],
    # Nombre de las columnas de encabezado que tendra el archivo de salida TXT
    'ENCABEZADO_TXT': ['CRR', 'CALIDAD', 'RUT'],
    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'RUT': 1, 'CALIDAD': 3 }
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
    'ENCABEZADO_TXT': ['CRR', 'COBRANZA_PRO', 'PACPAT_PRO', 'ESTADO_PRO', 'ESTADO_UT_PRO', 'RUT', 'ID_CAMPANA', 'CAMPANA', 'POLIZA', 'ID_CLIENTE'],

    # Nombre de las columnas de encabezado que tendra el archivo de POLIZAS RELOQUIDADAS
    'ENCABEZADO_RELIQUIDACIONES': ['CRR', 'COBRANZA_REL_PRO', 'PACPAT_REL_PRO', 'ESTADO_PRO', 'ESTADO_UT_PRO', 'RUT', 'ID_CAMPANA', 'CAMPANA', 'POLIZA', 'ID_CLIENTE'],

    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['NOMBRE', 'FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'CAMPAIGN MEMBER: ID EMPLEADO', 'ESTADO', 'FECHA DE CIERRE', 'PÓLIZA: NUMERO DE PÓLIZA', 'FECHA DE EXPIRACIÓN DEL CO-RET', 'ESTADO DE RETENCIÓN', 'MIEMBRO DE CAMPAÑA ID.', 'ESTADO DE ÚLTIMA TAREA'],

    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'NOMBRE_CLIENTE': 0, 'FECHA_CREACION': 1, 'NOMBRE_DE_CAMPAÑA': 2, 'CODIGO_EJECUTIVO': 3, 'ESTADO': 4, 'FECHA_CIERRE': 5, 'NRO_POLIZA': 6, 'EXPIRACION_CORET': 7, 'ESTADO_RETENCION': 8, 'CAMAPAÑA_ID': 9, 'ESTADO_ULTIMA_TAREA': 10},

    # Valores para los Estado de Última Tarea
    'LISTA_ULTIMA_TAREA' : {'Numero invalido': 1, 'Cliente retenido': 2, 'Llamado reprogramado': 3, 'Cliente no retenido': 4, 'Sin respuesta': 5, 'Buzón de voz': 6, 'Pagos al día': 7, 'Teléfono ocupado': 8, 'Teléfono apagado':	9, 'No quiere escuchar': 10, 'Contacto con el asesor': 11, 'Campaña completada con 5 intentos': 12, 'Apoyo del asesor al ejecutivo': 13, 'Número equivocado': 14, 'Pendiente respuesta cliente': 15, 'Sin gestión de cierre': 16, 'Sin teléfono registrado': 17, 'Cliente no actualizado': 18, 'Temporalmente fuera de servicio': 19, 'Carta de revocación pendiente': 20, 'Contacto por correo': 21, 'Campaña exitosa': 22, 'Solicita renuncia': 23, 'Cliente desconoce venta': 24, 'No se pudo instalar mandato':	25, 'Anulado por cambio de producto Metlife': 26, 'Cliente vive en el extranjero': 27, 'Cliente activa mandato': 28, 'Plazo previsto del producto': 28, 'Queda vigente sin pagar': 30, 'Lo está viendo con Asesor': 31},
}

REACTIVA_CONFIG_XLSX = {
    # Nombre que tendra el proceso dentro del flujo
    'PROCESO': 'REACTIVA',
    # Coordenadas encabezado
    'COORDENADA_ENCABEZADO': 'A1:N1',
    # Argumentos que necesita el proceso para funcionar
    'ARGUMENTOS_PROCESO': 4,
    # Nombre del archivo TXT que el proceso generara
    'SALIDA_TXT': 'REACTIVA',

    # Nombre del archivo XLSX que el proceso usara
    'ENTRADA_XLSX': '',

    # Nombre del archivo base de certificacion
    'ARCHIVO_BASE_CERTIFICACION': {

        # Archivo de entrada
        'NOMBRE_ARCHIVO': 'Base Certificacion',

        # Nombre de las columnas del encabezado que tendra el archivo de Base Certificacion
        'ENCABEZADO': ['PÓLIZA', 'REQUERIMIENTO', 'FECHA DE LLAMADO', 'HORA DE LLAMADO', 'ID EMPLEADO', 'CANAL', 'SE ENCONTRÓ GRABACIÓN EN NICE', 'CLIENTE CONTACTADO', 'UTILIZA ARGUMENTO DE DESACTIVACIÓN DE MEDIO DE PAGO', 'CONFIRMA NOMBRE COMPLETO DE CLIENTE ', 'CONFIRMA RUT DEL CLIENTE', 'CONFIRMA TELÉFONO DE CONTACTO', 'EJECUTIVA DEJA SIN EFECTO CARTA EN FORMA VOLUNTARIA CLIENTE', 'MENCIONA NÚMERO DE PÓLIZA Y NOMBRE DEL PRODUCTO', 'MENCIONA QUE MANTENDRÁ COBERTURA VIGENTE', 'ENVÍA CARTA DE REVOCACIÓN ', 'EXPLICA COMO COMPLETAR CARTA DE REVOCACIÓN', 'TIPO DE CERTIFICACIÓN', 'FECHA CERTIFICACIÓN', 'EMAIL ENVIADO', 'GESTIÓN', 'UF', 'CARTA', 'EXISTE MIEMBRO DE CAMPAÑA - TERMINADO CON ÉXITO EN SALESFORCE'],

        # # Columnas que se utilizaran durante el procesamiento del archivo Base de Certificacion
        'COLUMNAS': {'NRO_POLIZA': 0, 'FECHA_LLAMADO': 2, 'EJECUTIVO': 4, 'CANAL': 5, 'TIPO_CERTIFICACION': 17}
    },

    # Nombre de las columnas del encabezado que tendra el archivo (se usa para validar que el archivo este correcto)
    'ENCABEZADO_XLSX': ['FECHA DE CREACIÓN', 'CAMPAÑAS: NOMBRE DE LA CAMPAÑA', 'ESTADO', 'PÓLIZAS EN CAMPAÑA', 'FECHA DE CIERRE', 'PÓLIZA: NUMERO DE PÓLIZA', 'ESTADO DE RETENCIÓN', 'ÚLTIMA FECHA DE ACTIVIDAD', 'MIEMBRO DE CAMPAÑA ID.', 'ES LLAMADA SALIENTE', 'CAMPAIGN MEMBER: ID EMPLEADO'],

    # Columnas que se utilizaran durante el procesamiento del archivo XLSX
    'COLUMNAS_PROCESO_XLSX': {'FECHA_CREACION': 0, 'ESTADO': 2, 'FECHA_CIERRE': 4, 'NRO_POLIZA': 5, 'ESTADO_RETENCION': 6, 'CAMAPAÑA_ID': 8, 'ESTADO_ULTIMA_TAREA': 9, 'LLAMADA_SALIENTE': 10, 'NOMBRE_EJECUTIVO': 11},

    # Estados de Retencion
    'ESTADOS_RETENCION': {'Cliente no vigente': 1, 'Cliente al día': 2, 'Lo va a pensar': 3, 'Solicita contacto por correo': 4, 'Pendiente de endoso': 5, 'Mantiene su producto': 6, 'Desiste el producto': 7, 'Anulado por cambio de producto': 8, 'Sin gestión por cierre': 9, 'Término programado de producto': 10, 'Queda vigente sin pagar': 11, 'Espera de carta de anulación': 12},

    # Configuracion de archivos de salida
    'SALIDA_TXT':{
        'GESTION':{
            'NOMBRE_SALIDA': 'GESTION_REACT',
            'ENCABEZADO': 'ESTADO_VALIDO_REACT;CONTACTO_REACT;EXITO_REPETIDO_REACT;RUT;ID_CAMPANA;CAMPANA;POLIZA'
        },
        'POLIZA':{
            'NOMBRE_SALIDA': 'POLIZA_REACT',
            'ENCABEZADO': 'ESTADO_POLIZA_REACT;POLIZA'
        },
        'CERTIFICACION':{
            'NOMBRE_SALIDA': 'CERTIFICACION_REACT',
            'ENCABEZADO': 'GRAB_CERTIFICADA_REACT;RUT;CAMPANA;POLIZA'
        },
    }
}

COMPLEMENTO_CLIENTE_XLSX = {
    # Nombre del archivo Complemento Cliente
    'NOMBRE_ARCHIVO': 'COMPLEMENTO CLIENT vLite',

    # Nombre de las columnas del encabezado que tendra el archivo de Complemento Cliente
    'ENCABEZADO': ['NROPOLIZA', 'NROCERT', 'ESTADOPOLIZA', 'FEC_ULT_PAG', 'ESTADO_MANDATO', 'FECHA_MANDATO', 'FECHAPROCESO'],

    # # Columnas que se utilizaran durante el procesamiento del archivo Complemento Cliente
    'COLUMNAS': {'NRO_POLIZA': 0, 'NRO_CERT': 1, 'ESTADO_POLIZA': 2, 'FEC_ULT_PAG': 3, 'ESTADO_MANDATO': 4, 'FECHA_MANDATO': 5},
}