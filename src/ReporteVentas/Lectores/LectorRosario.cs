using ClosedXML.Excel;
using ReporteVentas.Modelo;
using static ReporteVentas.Lectores.ExcelUtil;

namespace ReporteVentas.Lectores;

/// <summary>Rosario: una hoja por mes, títulos arriba, filas vacías en el medio y fila TOTAL al final.</summary>
public class LectorRosario : ILectorSucursal
{
    public string Sucursal => "Rosario";

    public bool PuedeLeer(string ruta) =>
        Path.GetExtension(ruta).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
        && Path.GetFileName(ruta).StartsWith("Rosario", StringComparison.OrdinalIgnoreCase);

    public IReadOnlyList<LecturaArchivo> Leer(string ruta, Catalogo catalogo)
    {
        var lecturas = new List<LecturaArchivo>();
        using var wb = new XLWorkbook(ruta);
        foreach (var ws in wb.Worksheets)
        {
            var lectura = new LecturaArchivo { Archivo = $"{Path.GetFileName(ruta)} › {ws.Name}", Sucursal = Sucursal };

            // El encabezado no está siempre en la misma fila: se busca la que dice "Fecha".
            int filaEncabezado = Enumerable.Range(1, 15).FirstOrDefault(i => ws.Row(i).CellsUsed().Any(c => TieneTexto(c, "Fecha")));
            if (filaEncabezado == 0) throw new InvalidDataException($"La hoja '{ws.Name}' no tiene fila de encabezado.");
            lectura.FilasSalteadas += filaEncabezado; // títulos, filas vacías de arriba y encabezado

            var h = Encabezados(ws.Row(filaEncabezado));
            int cFecha = Col(h, false, "fecha"), cCod = Col(h, true, "cod", "codigo"), cProd = Col(h, false, "descripcion", "producto"),
                cCat = Col(h, true, "rubro", "categoria"), cCant = Col(h, false, "unidades", "cantidad"),
                cPrecio = Col(h, true, "precio"), cImporte = Col(h, false, "importe"), cVend = Col(h, false, "vendedor");

            for (int r = filaEncabezado + 1; r <= ws.LastRowUsed()!.RowNumber(); r++)
            {
                var fila = ws.Row(r);
                if (fila.IsEmpty()) { lectura.FilasSalteadas++; continue; }
                if (fila.CellsUsed().Any(c => TieneTexto(c, "TOTAL")))
                {
                    // No se suma, pero se guarda para cruzarlo contra lo calculado.
                    lectura.TotalDeclarado = (lectura.TotalDeclarado ?? 0) + Numero(fila.Cell(cImporte));
                    lectura.FilasSalteadas++;
                    continue;
                }
                lectura.FilasLeidas++;
                ConstructorVenta.Agregar(lectura, catalogo, Sucursal, Fecha(fila.Cell(cFecha)),
                    cCod > 0 ? fila.Cell(cCod).GetString() : null, fila.Cell(cProd).GetString(),
                    cCat > 0 ? fila.Cell(cCat).GetString() : null, (int)Numero(fila.Cell(cCant)),
                    cPrecio > 0 ? Numero(fila.Cell(cPrecio)) : null, Numero(fila.Cell(cImporte)), fila.Cell(cVend).GetString());
            }
            lecturas.Add(lectura);
        }
        return lecturas;
    }
}
