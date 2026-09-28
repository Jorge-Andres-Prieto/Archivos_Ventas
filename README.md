# Procesamiento de Ventas

Automatización en Python que toma un archivo de ventas, lo procesa y guarda los resultados en una carpeta específica, dejando un registro de cada ejecución.

## Más allá de lo solicitado

Además de los requerimientos del ejercicio, se tuvieron en cuenta dos aspectos pensando en un caso real:

- **Registro de logs:** cada ejecución guarda la fecha, la hora y si terminó con éxito o con error.
- **Organización por carpetas:** cada tipo de archivo tiene su lugar, lo que hace el proceso reutilizable y automatizable. Un flujo simple: *se toma un archivo → se procesa → se guarda en su carpeta*.

## Estructura del proyecto

```
Archivos_Ventas/
├── Datos Ventas/          Archivo original
├── Resumen de Ventas/     Archivos generados
├── Logs/                  Registro de ejecuciones
├── procesamiento_datos.py Lógica general
└── actualizar_logs.py     Manejo de logs
```

## Archivos

**Datos Ventas**
Contiene `datos_ventas.xlsx`, el archivo original. No se modifica.

**Resumen de Ventas**
Contiene los archivos generados:

- `datos_ventas_actualizado.xlsx`: copia del original con los cambios aplicados (fechas convertidas a formato fecha, totales faltantes calculados y una nueva columna `Mes`).
- `resumen_ventas.xlsx`: el resumen solicitado, con dos hojas:
  - **Resumen_Ventas:** total de ventas de 2023 por vendedor.
  - **Ventas_Mensuales:** total de ventas de 2023 por mes.

**Logs**
Contiene `logs_ejecucion.xlsx`, donde cada ejecución agrega una fila con fecha, hora, estado (`EXITOSO` o `ERROR`) y detalle.

## Código

El proyecto usa dos archivos de Python:

- `procesamiento_datos.py`: maneja toda la lógica general (leer, procesar y guardar).
- `actualizar_logs.py`: maneja únicamente los logs.

Separar las responsabilidades hace que el código sea más **modular y fácil de mantener**: cualquier cambio o mejora se hace en un solo lugar sin afectar lo demás.

## ¿Qué hace el programa?

1. Lee el archivo original de ventas.
2. Completa los totales faltantes (`Cantidad × Precio_Unitario`).
3. Convierte la fecha y agrega la columna `Mes`.
4. Guarda el archivo actualizado.
5. Genera el resumen de ventas de 2023 con formato visual.
6. Registra el resultado de la ejecución en el log.

## Uso

```
pip install pandas openpyxl
python procesamiento_datos.py
```
