# evaluacion_ceapsi
Herramienta para tabulación automática de datos para el departamento de evaluación psicológica del CEAPSI UAGRM

- El **Archivo de datos** es el archivo que sacan del formulario de google.
- La **Carpeta destino** es la carpeta en la que se guardarán los archivos individuales por estudiante.
- La **Fila inicial** es la fila desde la que se procesará el archivo de datos.
- La **Fila final** es la fila hasta la que se procesará el archivo de datos, debe ser mayor o igual a la fila inicial.
- La **Configuración** está compuesta del archivo plantilla y el archivo de mapeo:

### Archivo plantilla
Es el archivo excel con macros en el que se encuentran sus fórmulas

### Archivo de mapeo
Es el archivo que le indica al programa a qué hoja, fila y columna se debe copiar desde el archivo de datos. 
Debe ser un archivo CSV delimitado por comas ",", el cuál puede ser editado y creado en excel.
El formato de cada fila en este archivo es el siguiente:

<columna_en_archivo_de_datos>,<nro_hoja_archivo_plantilla>,<nro_fila_archivo_plantilla>,<nro_columna_archivo_plantilla>,<_formula>

Por ejemplo, si quiero que la columna 2 del archivo de datos vaya a la hoja 1, en la fila 8 y la columna "C", sería lo siguiente:

Sin fórmula:

8,1,23,6

Con fórmula:

8,1,23,6,"=IF(""?""=""Hombre"", ""M"", ""F"")"

### Editar el archivo de mapeo
Para editar el archivo la recomendación es realizar lo siguiente:
- Abrir un libro nuevo de Excel.
- En la pestaña "Datos", sección "Origen de datos", seleccionar "Desde texto".
- En la ventana emergente seleccionar el archivo de mapeo.
- Click en siguiente
- En la siguiente ventana, seleccione la casilla de coma "," y click en siguiente
- En la siguiente ventana, seleccione la última columna y seleccione "Texto" como tipo de dato, luego click  en aceptar.


Con esto deberia poder visualizar los datos, donde puede proceder a  realizar sus cambios. Una vez terminados los cambios, para guardar realicelo de la siguiente forma:
- Ir a archivo y selecionar la opción "Guardar como".
- Abajo del nombre en la lista desplegable seleccione el tipo de archivo como "csv delimitado por comas".
- Click en guardar.
- Aceptar las ventanas emergentes que puedan salir.


Para referencia o más información pueden consultar el  siguiente artículo https://support.microsoft.com/es-es/office/importar-o-exportar-archivos-de-texto-txt-o-csv-5250ac4c-663c-47ce-937b-339e391393ba
