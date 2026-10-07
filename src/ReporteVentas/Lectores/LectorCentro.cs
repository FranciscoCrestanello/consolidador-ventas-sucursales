using ClosedXML.Excel;
using ReporteVentas.Modelo;
using static ReporteVentas.Lectores.ExcelUtil;

namespace ReporteVentas.Lectores;

/// <summary>Centro: formato "bueno" de referencia, sin columna Importe.</summary>
public class LectorCentro : ILectorSucursal
{
    public string Sucursal => "Centro";

    public bool PuedeLeer(string ruta) =>
        Path.GetExtension(ruta).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
        && Path.GetFileName(ruta).StartsWith("Ventas_Centro", StringComparison.OrdinalIgnoreCase);

    public IReadOnlyList<LecturaArchivo> Leer(string ruta, Catalogo catalogo)
    {
        var lectura = new LecturaArchivo { Archivo = Path.GetFileName(ruta), Sucursal = Sucursal };
        using var wb = new XLWorkbook(ruta);
        var ws = wb.Worksheets.First();
        var h = Encabezados(ws.Row(1));
        int cFecha = Col(h, false, "fecha"), cCod = Col(h, true, "codigo"), cProd = Col(h, false, "producto"),
            cCat = Col(h, true, "categoria"), cCant = Col(h, false, "cantidad"),
            cPrecio = Col(h, false, "preciounitario"), cVend = Col(h, false, "vendedor");
        lectura.FilasSalteadas++; // encabezado

        for (int r = 2; r <= ws.LastRowUsed()!.RowNumber(); r++)
        {
            var fila = ws.Row(r);
            if (fila.IsEmpty()) { lectura.FilasSalteadas++; continue; }
            lectura.FilasLeidas++;
            ConstructorVenta.Agregar(lectura, catalogo, Sucursal, Fecha(fila.Cell(cFecha)),
                cCod > 0 ? fila.Cell(cCod).GetString() : null, fila.Cell(cProd).GetString(),
                cCat > 0 ? fila.Cell(cCat).GetString() : null, (int)Numero(fila.Cell(cCant)),
                Numero(fila.Cell(cPrecio)), null, fila.Cell(cVend).GetString());
        }
        return [lectura];
    }
}
