# Reporte de ventas consolidado de 4 sucursales

> **El reporte que llevaba dos días ahora sale en segundos.**

**Distribuidora Del Sur** (empresa ficticia) tiene 4 sucursales. Cada mes, cada una manda sus ventas en un archivo distinto, con su propio formato. Hoy una persona de administración pasa casi dos días copiando, limpiando y juntando todo a mano.

Este programa se ejecuta con doble clic, lee todos los archivos de la carpeta `entrada`, los limpia, los junta y genera `salida/Reporte_Ventas.xlsx` con los números y gráficos listos.

<!-- Acá van las capturas del antes (carpeta con archivos desordenados) y del después (el reporte), y el video de 2 minutos. -->

## Qué resuelve

Cada archivo trae a propósito un problema distinto, los mismos que aparecen en clientes reales:

| Archivo | Sucursal | Problemas que resuelve el programa |
| --- | --- | --- |
| `Ventas_Centro_T3_2026.xlsx` | Centro | Formato "bueno" de referencia (sin columna de importe). |
| `norte_ventas_MM-2026.xlsx` (3) | Norte | Un archivo por mes, columnas con otros nombres, fecha guardada como texto, sin código ni categoría, productos en MAYÚSCULAS o con espacios de más. |
| `Rosario - Ventas 3er trimestre.xlsx` | Rosario | Una hoja por mes, títulos arriba, filas vacías en el medio y fila `TOTAL` que no hay que sumar. |
| `ventas_cordoba_2026T3.csv` | Córdoba | CSV con `;` y coma decimal, fecha `aaaa-mm-dd`, productos en minúsculas, importe total en vez de precio unitario y filas duplicadas. |
| `Catalogo_Productos.xlsx` | Todas | Catálogo maestro: completa código, nombre oficial y categoría. |

Todo termina en un formato único: `Fecha, Sucursal, Código, Producto, Categoría, Cantidad, Precio unitario, Importe, Vendedor`.

## El reporte

| Hoja | Contenido |
| --- | --- |
| **Resumen** | Período, ventas totales, operaciones, ticket promedio, mejor sucursal, mejor vendedor y producto más vendido. |
| **Por sucursal y mes** | Tabla de sucursales por meses con totales y gráfico de columnas. |
| **Productos** | Top 10 por importe, ventas por categoría y gráfico. |
| **Datos** | Todas las filas consolidadas, como tabla de Excel con filtros. |
| **Control** | Filas leídas por archivo, duplicados quitados, productos que no están en el catálogo y verificación contra los `TOTAL` que traen los archivos. |

La hoja **Control** es la que le da confianza al cliente: muestra exactamente qué encontró y qué descartó.

## Resultado

| Sucursal | Operaciones | Ventas (AR$) |
| --- | --- | --- |
| Centro | 691 | 5.762.270 |
| Norte | 435 | 3.482.100 |
| Rosario | 501 | 3.949.680 |
| Córdoba | 356 | 2.970.130 |
| **Total** | **1.983** | **16.164.180** |

Córdoba trae 361 filas; se quitan 5 duplicadas. Estos números están verificados por los tests.

## Cómo usarlo

1. Poné `ReporteVentas.exe` en una carpeta que tenga la subcarpeta `entrada/` con los archivos del mes.
2. Doble clic. Se crea `salida/Reporte_Ventas.xlsx` y se abre solo.

No necesita instalar nada: es un único `.exe` autocontenido.

## Cómo está hecho

- **C# / .NET 10**, aplicación de consola publicada como un solo `.exe`.
- **ClosedXML** y **CsvHelper** para leer; **EPPlus 4.5.3** (última versión gratuita) para escribir el reporte con tablas y gráficos nativos de Excel.
- **Un lector por sucursal** (`Lectores/`), todos devolviendo la misma clase `Venta`. **Sumar una sucursal nueva es agregar un lector** y registrarlo en `Consolidador.Lectores`.
- Las columnas se buscan por nombre de encabezado (sin importar tildes ni mayúsculas) y no por posición, así que un archivo con las columnas en otro orden sigue funcionando.
- Los duplicados se quitan solo donde corresponde (Córdoba). Las otras sucursales tienen ventas legítimas idénticas el mismo día, que no deben perderse.
- Si el reporte está abierto en Excel, avisa con un mensaje claro en vez de fallar.

```
src/ReporteVentas/     programa (Modelo, Lectores, Catalogo, Consolidador, Reporte)
tests/ReporteVentas.Tests/   tests xUnit contra la tabla de resultados esperados
entrada/               archivos de ejemplo
```

### Compilar y probar

```powershell
dotnet test
dotnet publish src/ReporteVentas -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o publicar
```

---

*Todo el código y los datos son propios y ficticios.*
