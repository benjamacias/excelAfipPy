# excelAfipPy
Este programa en Python automatiza la consolidación de archivos de comprobantes electrónicos generados por la AFIP en formato Excel o CSV. Su objetivo principal es analizar los valores económicos de cada tipo de comprobante, clasificados por columnas y volcar los resultados en una hoja resumen dentro del mismo archivo


Lectura automática de archivos
Escanea la carpeta recibidos/ y detecta todos los archivos .xlsx o .csv de comprobantes a procesar.

Cálculo de importes por tipo de comprobante
Para cada archivo, el programa agrupa y suma los valores por tipo de comprobante (por ejemplo, “Factura A”, “Factura B”, “Nota de Crédito B”, etc.) y por tipo de importe:Imp. Neto Gravado, imp. Neto No Gravado, IVA, Total

Escritura de resumen en Excel
Los totales calculados se escriben en un bloque específico del mismo archivo Excel (S17:Y21) generando así un resumen financiero claro y estandarizado.

Renombrado inteligente del archivo
Utiliza el CUIL detectado en el nombre del archivo y lo cruza con un archivo de clientes (clientes.xls) para obtener el nombre del cliente. El archivo se guarda renombrado en la carpeta terminado/ con el siguiente formato:


Libro venta CLIENTE MES-AÑO.xlsx
Cálculo del coeficiente de IVA
Al final de cada archivo, calcula el coeficiente de IVA del último comprobante válido y lo deja registrado como referencia adicional.

## Microservicio

El microservicio expone dos endpoints principales:

- `POST /process` procesa todos los archivos en la carpeta `recibidos/`.
- `POST /process-files` permite subir varios archivos XLSX y devuelve un ZIP con los archivos procesados usando `procesar_archivo`.

