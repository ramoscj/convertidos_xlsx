### Instalación de componentes necesarios:

+ Descargar e instalar una versión de [Python 3.7.X](https://www.python.org/downloads/  "Python 3.7.X"), marcar la opción para agregar al PATH la variable de Python al sistema.

![](https://i.postimg.cc/MG581vfz/pythonsetup-2.jpg)

  
+ Después de finalizada la instalación de Python abrir el CMD de Windows y para verificar que la instalación esta correcta y esta configurada la variable de Python en el sistema ingresa el siguiente comando "Python" en el CMD y este deberá mostrar el interprete de Python.


![](https://i.postimg.cc/gj6zBLhs/python1.png)


+ Escribir el comando "exit()" para salir y volver al cursor de la consola, una vez en el cursor de la consola dirigirse al directorio del archivo "requirements.txt" (esta dentro de los archivos de la aplicación) y escribir lo siguiente para instalar las librerías necesarias.


```bash
pip install -r requirements.txt
```

## Proceso DOTACIÓN

El archivo que controla la ejecución de este proceso es "procesoDotacion.py", para ejecutarlo se debe escribir el siguiente comando en la consola (CMD), el formato que debe tener el parámetro **FECHA** debe ser YYYMM.

**python procesoDotacion.py** + fecha del proceso + archivoAsistencia.xlsx + carpeta donde se generara la "salida.txt".

```python
python .\procesoDotacion.py 202006 .CRO\INPUTS\202006_Asistencia_Plataformas.xlsx .CRO\OUTPUTS
```

Este proceso genera la salida de los archivos:

+  **Archivo de Dotación:** ICOM_CA_MTLFCC_YYYYMM.txt.

+  **Archivo de Asistencia:** ASISTENCIAYYYYMM.txt.


## Proceso CRO

El archivo que controla la ejecución de este proceso es "procesoCro.py", para ejecutarlo se debe escribir el siguiente comando en la consola (CMD), el formato que debe tener el parámetro **FECHA** debe ser YYYMM y el rango de **FECHAS** debe ser YYYYMMDD.

**python procesoCro.py** + nombre del proceso + fecha del proceso + archivo data.xlsx + carpeta donde se generara la "salida.txt".

Este proceso solo tiene una variante para la ejecución de "Gestión" ya que dicho proceso necesita un rango de fechas adicional (formato de fecha YYYYMMDD) y un archivo adicional .xlsx que es el de propietarios.

**python procesoCro.py** + gestion + fecha del proceso + fecha rango inicio + fecha rango fin + archivoGestion.xlsx + archivoPropietarios.xlsx + carpeta donde se generara la "salida.txt".

```python
python .\crear_txt.py fuga 202011 .CRO\INPUTS\202006_Fuga_Agencia.xlsx .CRO\OUTPUTS
python .\crear_txt.py campanha_especial 202011 .CRO\INPUTS\202011_CampañasEspeciales_CRO.xlsx .CRO\OUTPUTS
python .\crear_txt.py calidad 202011 .CRO\INPUTS\202011_Calidad_CRO.xlsx .CRO\OUTPUTS
python .\crear_txt.py gestion 202011 20201026 20201123 '.CRO\INPUTS\Gestión CRO.xlsx .CRO\INPUTS\Propietarios CRO.xlsx .CRO\OUTPUTS
```

**Archivos que genera cada ejecución:**
+ **python procesoCro.py fuga**: Genera los archivos de FUGAYYYYMM.txt
+ **python procesoCro.py campanha_especial**: Genera el archivo de PILOTOYYYYMM.txt
+ **python procesoCro.py calidad**: Genera el archivo de CALIDADYYYYMM.txt
+ **python procesoCro.py gestion**: Genera el archivo de GESTIONYYYYMM.txt
 
 ### NOTA:
 
 Para el proceso de "FUGA" se debe considerar para su ejecución que el mismo genera la información para el periodo anterior según la  fecha de proceso que se esta ingresando como parámetro, es decir si se ingresa como periodo 202002 (Febrero 2020) el proceso restara un mes al periodo ingresado y quedaría 202001 (Enero 202001 para la ejecución interna del proceso), por tal motivo el archivo que se le pasa como parámetro para su lectura tiene que ser un periodo anterior al que se quiere procesar (para este ejemplo 202001_Fuga_Agencia.xlsx) y al finalizar el procesamiento del archivo generaría la salida para el periodo ingresado como parámetro (Febrero 2020).

## Proceso PROACTIVA

El archivo que controla la ejecución de este proceso es "procesoProactiva.py", para ejecutarlo se debe escribir el siguiente comando en la consola (CMD), el formato que debe tener el parámetro **FECHA** debe ser YYYMM.

**python procesoProactiva.py** + fecha del proceso + archivoGestionProactiva.xlsx + archivoComplementoCliente.xlsx + carpeta donde se generara la "salida.txt".

```python
python crearTxtProactiva.py 202102 .\PROACTIVA\INPUTS\202102_Gestion_CoRet_Proactiva.xlsx .\PROACTIVA\INPUTS\202102_Complemento_Cliente_Proactiva.xlsx .\PROACTIVA\OUTPUTS
```

Este proceso genera la salida de los archivos:

+  **Archivo de base proactiva:** ICOM_GESTION_CORPROYYYYMM.txt.

+  **Archivo de pólizas reliquidadas:** ICOM_RELIQUIDACION_CORPROYYYYMM.txt.

## Proceso REACTIVA

El archivo que controla la ejecución de este proceso es "procesoReactiva.py", para ejecutarlo se debe escribir el siguiente comando en la consola (CMD), el formato que debe tener el parámetro **FECHA** debe ser YYYMM y el rango de **FECHAS** debe ser YYYYMMDD.

**python procesoReactiva.py** + fecha del proceso + archivoGestionProactiva.xlsx + archivoComplementoCliente.xlsx + carpeta donde se generara la "salida.txt".

```python
python .\crearTxtReactiva.py 202105 20210428 20210526 .\REACTIVA\INPUTS\202105_Gestion_CoRet_Reactiva.xlsx .\REACTIVA\INPUTS\202105_Base_Certificacion_Reactiva.xlsx .\REACTIVA\INPUTS\202105_Complemento_Cliente_Reactiva.xlsx .\REACTIVA\OUTPUTS
```

Este proceso genera la salida de los archivos:

+  **Archivo de base reactiva:** GESTION_REACTYYYYMM.txt.

+  **Archivo de pólizas certificadas:** CERTIFICACION_REACTYYYYMM.txt.

+  **Archivo de pólizas vigentes:** POLIZA_REACTYYYYMM.txt.


## Instalacion de la Base de Datos

El archivo con la base de datos esta en la carpeta "DB" hay se encuentra el .bak que es de un SQL Server. En el archivo **config_xlsx.py** existe una variable que contiene los parametros para conectarse a la DB que tiene por nombre **ACCESO_DB**, solo se deben cambiar los parametros de conexion.

> Si se va a utilizar el archivo .bacpac solo se debe cargar en la opción "Import Data-tier Aplication" con la herramienta SQL Server Management.


## NOTA

Para los fines de configuración el archivo **config_xlsx.py** contiene la variable que controla el directorio donde se encuentran los archivos de logs. Se debe asegurar que el usuario con que se ejecute el programa tenga permisos para escribir en estos directorios.

+  **PATH_LOG**: Directorio de archivos LOG de salida


### CRO

En el mismo archivo existen seis variables que tienen algunos parámetros para evaluar los archivos de entrada. Se dará ejemplo con la variable del archivo de FUGA y las demás también tendrán una estructura parecida:

+  **FUGA_CONFIG_XLSX:** Archivo de FUGA

+  **SALIDA_TXT:** Nombre del archivo de salida .TXT

+  **COORDENADA_ENCABEZADO:** Contiene las coordenadas del encabezado del archivo, para luego validarlo.

+  **ENCABEZADO_XLSX:** Contiene los nombre de las columnas que debe tener el archivo.

+  **ENCABEZADO_FUGA_TXT:** Formato del archivo de salida .txt.

+  **COLUMNAS_PROCESO_XLSX:** Numero de las columnas que se utilizan en el proceso.