# Consolidador de reportes Excel y CSV

Aplicación de escritorio para reunir reportes mensuales en un único Excel, en vista sábana o en hojas por período. Surgió de una necesidad concreta de un contador.

No calcula IVA, saldos ni asientos contables. Conserva los valores de las celdas y agrega trazabilidad y un resumen técnico. No requiere activación, conexión a Internet durante su uso ni una cuenta de usuario.

## Usar el ejecutable

Descargar y descomprimir el paquete de Windows, y abrir `ConsolidadorExcel.exe`. No requiere instalar Python. Mantener el ejecutable junto a su carpeta `_internal`.

## Ejecutar desde el código fuente

Requiere Python 3.12 o posterior con Tkinter. Las comprobaciones iniciales se hicieron con Python 3.14 en Windows. Desde una terminal abierta en esta carpeta:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gui.py
```

Si las dependencias ya están instaladas en tu Python, podés ejecutar `python gui.py` o abrir `INICIAR.cmd` con doble click. El lanzador no instala nada; muestra instrucciones si falta alguna dependencia.

## Cómo usarlo

1. Agregar archivos o una carpeta de entrada. La carpeta se lee sin recorrer subcarpetas y omite archivos temporales de Excel.
2. Elegir **Vista sábana** o **Hojas por período**.
3. Opcionalmente activar **Comprobar meses faltantes**, con Desde/Hasta en MM/AAAA. Avisa si falta Compras o Ventas por mes, sin filtrar los datos. No garantiza que estén todos los comprobantes. Presionar **Preparar y revisar**. Leer los avisos y la cantidad de filas.
4. Presionar **Guardar consolidado** y elegir un archivo `.xlsx`, preferentemente en otra carpeta.

El resumen de archivos y los avisos se muestran solo en pantalla: no se agregan hojas de revisión al Excel. Los controles de selección se bloquean durante el trabajo; la ventana debe permanecer abierta hasta terminar.

## Ejemplo reproducible

Seleccionar juntos `examples/compras_enero.xlsx` y `examples/ventas_febrero.csv`.

- Resultado esperado: **2 archivos y 4 filas de datos**.
- Vista sábana: 4 filas en una única hoja, sin hojas adicionales.
- Hojas por período: 2 filas en `2026-01` y 2 en `2026-02`.
- Los identificadores de texto mantienen sus ceros iniciales.
- Los importes de CSV permanecen como texto: no se convierten ni se suman.

Los archivos son enteramente ficticios. No contienen datos de clientes.

## Entradas admitidas y límites

- Excel `.xlsx` y `.xlsm`: exactamente una hoja con datos. Se leen valores, no se copian macros, estilos, gráficos ni formatos de visualización. Las fórmulas y errores de celda se rechazan para evitar resultados sin recalcular. Una macro nunca se ejecuta.
- Excel antiguo `.xls`: requiere `xlrd`, incluido en las dependencias. Este formato utiliza los valores almacenados en el libro; no recalcula fórmulas ni garantiza detectar su presencia. Usar reportes exportados como valores.
- CSV: UTF-8 (con o sin BOM) o Windows-1252; separado por coma, punto y coma o tabulación. Respeta campos entre comillas. El número de fila de origen es el número de registro CSV; un campo con saltos de línea puede ocupar varias líneas físicas.
- Se buscan **Fecha** o **Fecha de Emisión**, junto con **Tipo** o **Tipo de Comprobante** en los primeros 30 registros. Se aceptan espacios exteriores y diferencias de mayúsculas/acentos en la detección.
- Identifica Compras/Ventas por el nombre del archivo o, como respaldo, por encabezados Vendedor/Comprador. No cambia los campos originales. Si no puede identificarlo, avisa y lo ubica después de Compras/Ventas dentro del mes; no lo cuenta como cobertura por tipo.
- Fechas para agrupación: `DD/MM/AAAA`, `DD-MM-AAAA`, `AAAA-MM-DD`, `AAAA-MM-DD HH:MM:SS` o fechas nativas de Excel. Una fecha inválida o vacía se conserva, con aviso, en **Sin fecha válida**. No se convierte un número suelto en una fecha.
- Se omiten títulos previos a la cabecera, filas completamente vacías y columnas completamente vacías. Se informa la cantidad de filas vacías omitidas. Los encabezados pierden espacios exteriores; los valores de datos no se recortan.
- Columnas sin título pero con datos, encabezados duplicados y archivos con varias hojas de datos requieren corrección explícita. Si un archivo falla, se bloquea toda la preparación: no se exporta silenciosamente un subconjunto.
- Si los encabezados difieren, se conserva la unión de columnas y se avisa. No se fusionan automáticamente `Total` e `Imp. Total`, por ejemplo.
- Las filas repetidas se conservan. Seleccionar la misma ruta dos veces no duplica su contenido. Archivos distintos con igual nombre requieren renombrar una copia para identificar el origen.
- Se agregan `Origen - archivo`, `Origen - hoja`, `Origen - fila` y `Origen - período`. Estos nombres están reservados. No se admite reimportar una salida como si fuera un reporte original.
- La salida respeta los límites de filas, columnas y longitud de texto de Excel. Procesa los datos en memoria; no está diseñada ni medida para volúmenes masivos.
- Los ceros mostrados mediante formato numérico de Excel no son parte del valor almacenado. Para conservar identificadores con ceros iniciales, deben estar guardados como texto en la entrada. No recupera precisión ya perdida en el archivo fuente.
- La vista sábana ordena cronológicamente por mes, primero Compras y luego Ventas. Dentro de cada mes y tipo conserva el orden de selección y de filas. Las hojas por período reúnen Compras y Ventas en la misma hoja mensual.

## Protección de archivos

No permite usar una entrada como destino. Reemplazar un archivo existente requiere confirmación. Escribe primero un archivo temporal en la carpeta de destino y lo reemplaza solo al completar la exportación. Si Excel mantiene el destino bloqueado, se informa el error. El texto que comienza con `=` se guarda como texto literal, no como fórmula.

## Arquitectura

```text
gui.py                                  Punto de entrada
app/shell.py                            Interfaz y trabajo en segundo plano
modules/afip_consolidator/controller.py  Estado y coordinación
modules/afip_consolidator/importers.py   Lectura y validación técnica
modules/afip_consolidator/processor.py   Unión de datos
modules/afip_consolidator/exporter.py    Escritura segura y formato Excel
modules/afip_consolidator/models.py      Modelos compartidos
tests/                                  Pruebas con archivos sintéticos
```

La interfaz usa una cola para recibir resultados de un trabajador, sin actualizar widgets desde ese hilo. Cambiar la selección o repetir la preparación invalida el resultado anterior. El procesamiento se puede probar sin abrir la interfaz.

## Pruebas

```powershell
python -B -m unittest discover -v
```

Cubren conservación de filas y valores, ambas salidas, separadores y codificaciones CSV, fechas inválidas, identificadores, encabezados distintos, archivos incorrectos, protección de entradas y guardado fallido. Son 20 pruebas automatizadas. También se verificó la conservación de valores de un CSV nativo de AFIP de 9.727 filas en ambas salidas. Los datos reales no se incluyen en el repositorio.

## Desarrollo

Proyecto desarrollado a partir de una necesidad de un contador. Mi participación incluyó el relevamiento de requisitos, la definición de funcionalidades, las pruebas de uso y la revisión de resultados. Utilicé Codex como asistencia para generar y modificar código, refactorizar la aplicación y agregar pruebas automatizadas.

## Crear el ejecutable

En Windows, con Python 3.14:

```powershell
python -m pip install -r requirements-build.txt
python -m PyInstaller --noconfirm --clean --onedir --windowed --name ConsolidadorExcel --collect-all customtkinter --hidden-import xlrd --exclude-module scipy --exclude-module matplotlib --exclude-module IPython --exclude-module pytest --exclude-module sympy --exclude-module numba --exclude-module tables --exclude-module pyarrow gui.py
```

También se puede usar `COMPILAR.cmd` después de instalar las dependencias. El resultado se genera en `dist/ConsolidadorExcel/ConsolidadorExcel.exe`. Los ejemplos y las pruebas pertenecen al repositorio, no se incorporan al ejecutable.

Para comprobar el ejecutable con datos ficticios (crea archivos en una carpeta nueva):

```powershell
.\dist\ConsolidadorExcel\ConsolidadorExcel.exe --self-test prueba-ejecutable
```

Al finalizar, `prueba-ejecutable/resultado.json` debe indicar `OK`. La prueba abre la interfaz oculta, importa CSV y verifica ambas exportaciones.
