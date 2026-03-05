# 📊 Importador Dinámico de Excel a DBF (Visual FoxPro)

## 📝 Descripción
Este script es una herramienta automatizada y reutilizable escrita en Visual FoxPro, diseñada para importar datos desde archivos Excel (`.XLS` formato XL8) hacia tablas nativas `.DBF`.

El sistema resuelve problemas comunes de importación como la generación de "columnas fantasma" por parte de Excel, desajustes en los nombres de las cabeceras y errores de tipos de datos, utilizando un puente SQL dinámico que mapea la información por posición estricta.

-------------------
## ✨ Características Principales

**Mapeo Dinámico (Cero Hardcoding):** Lee la estructura de la tabla destino y construye automáticamente la consulta SQL de inserción. No es necesario modificar el código si se añaden o quitan columnas en la base de datos.

**Filtro de Columnas Fantasma:** Ignora automáticamente las columnas vacías o con formato residual que Excel suele exportar al final del archivo.

**Conversión de Tipos Segura:** Al pasar los datos por un cursor intermedio, FoxPro castea de forma segura los valores de texto del Excel a los tipos nativos del DBF (Numérico, DateTime, Memo, etc.).

**Arquitectura Modular:** Empaquetado en un `PROCEDURE` que recibe parámetros, permitiendo su uso en múltiples módulos del sistema.

--------------
## ⚙️ Requisitos

- Visual FoxPro 9.0 (o compatible).

- Archivo de origen en formato Excel 97-2003 (`.XLS`).

- Base de datos destino abierta o accesible en la ruta de trabajo.

---------
## 🚀 Instrucciones de Uso

#### 1. Preparación del entorno y ejecución
Para utilizar la herramienta, simplemente configura tu entorno, llama a la función ImportarExcelDinamico pasándole los 3 parámetros requeridos y asegúrate de colocar un RETURN para separar la ejecución de la declaración de la función.

```
** Configuración inicial
OPEN DATABASE C:\Eikon_Vs\Eikon\Eikon_V7_QA\Exe\datasql.dbc SHARED
m_empresa = 01
CD C:\Eikon_Vs\Eikon\Eikon_V7_QA\

** Ejecución (Ruta Excel, Nombre de Hoja, Tabla Destino)
DO ImportarExcelDinamico WITH "C:\Users\Eikon\Downloads\Nomina\NOMINA16.XLS", "frhwd010", "frhwd010"

** Obligatorio para detener el flujo principal
RETURN
```

--------------------
#### 2. Parámetros de la Función

| Parámetro | Tipo | Descripción | Ejemplo |
| :--- | :--- | :--- | :--- |
| lcRutaExcel | Character | Ruta absoluta al archivo .XLS. | "C:\ruta\archivo.XLS" |
| lcHojaExcel | Character | Nombre exacto de la pestaña dentro del Excel. | "frhwd010" |
| lcTablaDestino | Character | Nombre de la tabla física .DBF destino. | "frhwd010" |

---------
## 🧪 Notas para el equipo de QA

Al realizar pruebas de regresión o validación de carga con este script, se recomienda verificar lo siguiente:

**Validación de Filas:** Compara el RECCOUNT() de la tabla destino con el número de filas reales (sin encabezado) en el Excel original.

**Validación de Columnas:** Si el Excel presenta errores, revisa el FCOUNT() del alias temporal que se genera durante el IMPORT para detectar columnas basura.

**Campos Memo:** Verifica visualmente (mediante BROWSE) que los datos no se hayan desplazado. El puente SQL garantiza la integridad, pero siempre es buena práctica revisar la última columna importada.

## 👨‍💻 Mantenimiento y Autoría

- **Desarrollado por:** Yarbis (Desarrollo / QA)
- **Entorno de Pruebas:** Eikon_V7_QA