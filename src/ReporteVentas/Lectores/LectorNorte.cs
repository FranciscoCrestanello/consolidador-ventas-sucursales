using ClosedXML.Excel;
using ReporteVentas.Modelo;
using static ReporteVentas.Lectores.ExcelUtil;

namespace ReporteVentas.Lectores;

/// <summary>Norte: un archivo por mes, columnas con otros nombres, fecha como texto, sin código ni categoría.</summary>
public class LectorNorte : ILectorSucursal
{
    public string Sucursal => "Norte";

    public bool PuedeLeer(string ruta) =>
        Path.GetExtension(ruta).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
        && Path.GetFileName(ruta).StartsWith("norte_ventas_", StringComparison.OrdinalIgnoreCase);

    public IReadOnlyList<LecturaArchivo> Leer(string ruta, Catalogo catalogo)
    {
        var lectura = new LecturaArchivo { Archivo = Path.GetFileName(ruta), Sucursal = Sucursal };
        using var wb = new XLWorkbook(ruta);
        var ws = wb.Worksheets.First();
        var h = Encabezados(ws.Row(1));
        int cFecha = Col(h, false, "fechaventa", "fecha"), cProd = Col(h, false, "articulo", "producto"),
            cCant = Col(h, false, "cant", "cantidad"), cPrecio = Col(h, false, "punit", "preciounitario", "precio"),
            cVend = Col(h, false, "vendedora", "vendedor");
        lectura.FilasSalteadas++; // encabezado

        for (int r = 2; r <= ws.LastRowUsed()!.RowNumber(); r++)
        {
            var fila = ws.Row(r);
            if (fila.IsEmpty()) { lectura.FilasSalteadas++; continue; }
            lectura.FilasLeidas++;
            ConstructorVenta.Agregar(lectura, catalogo, Sucursal, Fecha(fila.Cell(cFecha)),
                null, fila.Cell(cProd).GetString(), null, (int)Numero(fila.Cell(cCant)),
                Numero(fila.Cell(cPrecio)), null, fila.Cell(cVend).GetString());
        }
        return [lectura];
    }
}
