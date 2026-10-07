using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using ReporteVentas.Modelo;
using static ReporteVentas.Lectores.ExcelUtil;

namespace ReporteVentas.Lectores;

/// <summary>Córdoba: CSV con ';', coma decimal, importe total en vez de precio unitario y filas duplicadas.</summary>
public class LectorCordoba : ILectorSucursal
{
    public string Sucursal => "Córdoba";

    public bool PuedeLeer(string ruta) =>
        Path.GetExtension(ruta).Equals(".csv", StringComparison.OrdinalIgnoreCase)
        && Path.GetFileName(ruta).Contains("cordoba", StringComparison.OrdinalIgnoreCase);

    public IReadOnlyList<LecturaArchivo> Leer(string ruta, Catalogo catalogo)
    {
        var lectura = new LecturaArchivo { Archivo = Path.GetFileName(ruta), Sucursal = Sucursal };
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ";",
            PrepareHeaderForMatch = a => Clave(a.Header),
        };
        using var csv = new CsvReader(new StreamReader(ruta, System.Text.Encoding.UTF8, true), config);
        csv.Read();
        csv.ReadHeader();
        lectura.FilasSalteadas++; // encabezado

        var vistas = new HashSet<string>();
        while (csv.Read())
        {
            var campos = csv.Parser.Record!;
            if (campos.All(string.IsNullOrWhiteSpace)) { lectura.FilasSalteadas++; continue; }
            lectura.FilasLeidas++;
            // Duplicada = idéntica en todas sus columnas. Solo se aplica acá: en las otras sucursales
            // hay ventas legítimas que se repiten (mismo producto, cantidad y vendedor el mismo día).
            if (!vistas.Add(string.Join('\u001f', campos))) { lectura.DuplicadosQuitados++; continue; }

            ConstructorVenta.Agregar(lectura, catalogo, Sucursal, ParsearFecha(csv.GetField("fecha")!),
                csv.GetField("codigo"), csv.GetField("producto")!, null, int.Parse(csv.GetField("cantidad")!),
                null, NumeroEsAr(csv.GetField("importetotal")!), csv.GetField("vendedor")!);
        }
        return [lectura];
    }

    // "3640,00" o "1.234,50" → decimal, sin depender de la cultura de la computadora.
    private static decimal NumeroEsAr(string texto) =>
        decimal.Parse(texto.Trim().Replace(".", "").Replace(',', '.'), CultureInfo.InvariantCulture);
}
