 ### Instalación de componentes necesarios:

+ Descargar e instalar una version de [Python 3.7.X](https://www.python.org/downloads/ "Python 3.7.X"), marcar la opción para agregar al PATH la varible de python al sistema.

![](https://i.postimg.cc/MG581vfz/pythonsetup-2.jpg)

+ Despues de finalizada la instalacion de Python abrir el CMD de Windows y para verificar que la instalacion esta correcta y esta configurada la variable de Python en el sistema ingresa el siguiente comando "python" en el CMD y este debera mostrar el interprete de python.

![](https://i.postimg.cc/gj6zBLhs/python1.png)

+ Escribir el comando "exit()" para salir y volver al cursor de la consola, una vez en el cursor de la consola dirigirse al directorio del archivo "requirements.txt" (esta dentro de los archivos de la aplicación) y escribir lo siguiente para instalar las librerias necesarias.

```bash
pip install -r requirements.txt
```

## Proceso CRO
El archivo que controla la ejecucion es "crear_txt.py" y para ejecutarlo se debe escribir el siguiente comando en la consola (CMD).

**python crear_txt.py** + el proceso con sus respectivos parametros de entrada
```python
python crear_txt.py fuga 202009
python crear_txt.py asistencia 202009
python crear_txt.py campanha_especial 202009
python crear_txt.py calidad 202009
python crear_txt.py gestion 202009 20200901 20200930
```
El formato que debe tener el parametro **FECHA** debe ser YYYMM. La unica variante es para generar el archivo de **GESTION** que se deben ingresar tres **FECHAS** la primera es la del periodo a procesar (YYYYMM) y las otras dos son el rango de fecha del periodo (YYYYMMDD)

+ **python crear_txt.py fuga 202009**: Genera los archivos de FUGA
+ **python crear_txt.py asistencia 202009**: Genera los archivos de ASISTENCIA y DOTACION
+ **python crear_txt.py campanha_especial 202009**: Genera el archivo de PILOTO
+ **python crear_txt.py calidad 202009**: Genera el archivo de CALIDAD
+ **python crear_txt.py gestion 202009 20200901 20200930**: Genera el archivo de GESTION

## Proceso PROACTIVA
El archivo que controla la ejecucion es "crearTxtProactiva.py" y para ejecutarlo se debe escribir el siguiente comando en la consola (CMD). El formato que debe tener el parametro **FECHA** debe ser YYYMM.

**python crearTxtProactiva.py** + fecha del proceso + archivo PROACTIVA.xlsx + archivo COMPLEMENTO_CLIENTE.xlsx + carpeta donde se generara la "salida.txt".

```python
python crearTxtProactiva.py 202012 '.\PROACTIVA\INPUTS\Gestión CoRet Proactiva_202012.xlsx' '.\COMPLEMENTO_CLIENTE\COMPLEMENTO CLIENT vLite 202012.xlsx' .\PROACTIVA\OUTPUTS
```
> **NOTA**: Si los nombre de los archivos contienen espacios se deben ingresar entre comillas simples ('') y el directorio donde se genera la salida.txt no debe llevar el ultimo "\" ya que esta predeterminado en la configuracion.

## Instalacion de la Base de Datos

El archivo con la base de datos esta en la carpeta "DB" hay se encuentra el .bak que es de un SQL Server. En el archivo **config_xlsx.py** existe una variable que contiene los parametros para conectarse a la DB que tiene por nombre **ACCESO_DB**, solo se deben cambiar los parametros de conexion. 
> Si se va a utilizar el archivo .bacpac solo se debe cargar en la opción "Import Data-tier Aplication" con la herramienta SQL Server Management. 

## NOTA

Para los fines de configuracion el archivo **config_xlsx.py**  contiene las varibles que controlan los directorios donde se encuentran los archivos base, las salidas generadas y los logs. Se debe asegurar que el usuario con que se ejecute el programa tenga permisos para escribir en estos directorios.
+ **PATH_XLSX**: Directorio de archivos de entrada XLSX
+ **PATH_TXT**: Directorio de archivos de salida TXT
+ **PATH_LOG**: Directorio de archivos LOG de salida

### CRO
En el mismo archivo existen siete variables que tienen algunos parametros para evaluar los archivos de entrada. Se dara ejemplo con la variable del archivo de FUGA y las demas tambien tendran una estructura parecida:
 + **FUGA_CONFIG_XLSX**: Archivo de FUGA
   + **ENTRADA_XLSX**: Nombre del archivo de entrada .XLSX
   + **SALIDA_TXT**: Nombre del archivo de salida .TXT
   + **ENCABEZADO_XLSX**: Encabezado que debe tener el archivo de entrada .XLSX
   + **ENCABEZADO_FUGA_TXT**: Encabezado del archivo de salida .TXT
 + **ASISTENCIA_CONFIG_XLSX**: Archivo de ASISTENCIA
 + **GESTION_CONFIG_XLSX**: Archivo de GESTION
 + **CAMPANHAS_CONFIG_XLSX**: Archivo PILOTO
 + **CALIDAD_CONFIG_XLSX**: Archivo de CALIDAD
 + **DOTACION_CONFIG_XLSX**: Archivo de DOTACION
![](https://i.postimg.cc/gJkn1zSN/Captura.jpg)

## PROACTIVA
Para el proceso de "Proactiva" se crearon directorios apartes para no mezclar las salidas y entradas de los proceso, se encuentra en el mismo directorio que CRO solo que tendra una carpeta aparte con el nombre de "PROACTIVA" y tendra la misma esctructura.
 + **PROACTIVA**
   + **PROCESO_LOG**
   + **OUTPUTS**
   + **INPUTS**
