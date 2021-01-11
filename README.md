# evaluacion_ceapsi
Herramienta para tabulación automática de datos para el departamento de evaluación psicológica del CEAPSI UAGRM

-El *Archivo de datos* es el archivo que sacan del formulario de google.
-La *Carpeta destino* es la carpeta en la que se guardarán los archivos individuales por estudiante.
-La *Fila inicial* es la fila desde la que se procesará el archivo de datos.
-La *Fila final* es la fila hasta la que se procesará el archivo de datos, debe ser mayor o igual a la fila inicial.
-La *Configuración* está compuesta del archivo plantilla y el archivo de mapeo:

*Archivo plantilla*
Es el archivo excel con macros en el que se encuentran sus fórmulas

*Archivo de mapeo*
Es el archivo que le indica al programa a qué hoja, fila y columna se debe copiar desde el archivo de datos. 
Debe ser un archivo CSV delimitado por pipelines "|", el cuál puede ser editado y creado en excel.
El formato de cada fila en este archivo es el siguiente:
<columna_en_archivo_de_datos>|<nro_hoja_archivo_plantilla>|<nro_fila_archivo_plantilla>|<nro_columna_archivo_plantilla

Por ejemplo, si quiero que la columna 2 del archivo de datos vaya a la hoja 1, en la fila 8 y la columna "C", sería lo siguiente:
2|1|8|3